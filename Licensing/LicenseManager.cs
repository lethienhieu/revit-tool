using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json; // Thay cho System.Web.Script.Serialization
using System.Text.RegularExpressions;

namespace THBIM.Licensing
{
    public static class LicenseManager
    {
        // ============= CONFIG =============

        private static readonly string AppFolder =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "THBIM", "Licensing");

        private const string LegacyCacheFile = "thtools_suite.lic";
        private const string SessionFile = "session_v2_5.lic";
        private const int VerifyIntervalDays = 365;
        private const string VerifyStampFile = "verify.stamp";
        private const string LicenseServer = "https://script.google.com/macros/s/AKfycbxbMI884nGWzL6u7p10ChLqzOHSl4KXpJvr9EcsYP0lMiE3_ZNchDmNsDtLhOLEtKiYCw/exec";

        private static string Product = "TH-SUITE";
        public static void SetProduct(string productId)
        {
            if (!string.IsNullOrWhiteSpace(productId)) Product = productId.Trim();
        }

        private static string _lastError = null;

        static LicenseManager()
        {
            try
            {
                var asm = typeof(LicenseManager).Assembly;
                Log("=== THBIM.Licensing v2.5 loaded (Revit 2025/.NET 8) ===");
                Log("Assembly location = " + asm.Location);
                Log("Assembly version  = " + (asm.GetName().Version?.ToString() ?? "n/a"));
                Directory.CreateDirectory(AppFolder);
            }
            catch { }
        }

        // ============= PUBLIC API =============

        public static bool EnsureActivated()
        {
            Log("EnsureActivated (v2.5+gate)");
            try { Directory.CreateDirectory(AppFolder); } catch { }

            if (TryReadSession(out var ses) && !string.IsNullOrWhiteSpace(ses.Token))
            {
                try
                {
                    if (NeedVerify())
                    {
                        if (GetProfile(ses.Token, out var p))
                        {
                            ses.Email = p.Email ?? ses.Email;
                            ses.FullName = p.FullName ?? ses.FullName;
                            ses.Tier = string.IsNullOrWhiteSpace(p.Tier) ? ses.Tier : p.Tier.ToUpperInvariant();
                            ses.PremiumExpYMD = p.PremiumExpYMD ?? ses.PremiumExpYMD;
                            WriteSession(ses);
                            UpdateVerifyStamp();
                            Log($"VERIFY ok: {ses.Email}, tier={ses.Tier}, exp={ses.PremiumExpYMD}");
                        }
                        else
                        {
                            Log("VERIFY failed: token invalid or expired.");
                            // FIX: Thêm tiêu đề "THBIM"
                            Interaction.MsgBox("Please sign in using THBIM AutoUpdate app.", MsgBoxStyle.Information, "THBIM");
                            return false;
                        }
                    }
                }
                catch (Exception ex) { Log("VERIFY flow EX: " + ex); }

                var tier = (ses.Tier ?? "").Trim().ToUpperInvariant();
                var expDt = ParseDate(ses.PremiumExpYMD);
                var today = DateTime.UtcNow.Date;

                if (tier == "PREMIUM")
                {
                    if (expDt == null || expDt.Value.Date >= today)
                        return true;
                }

                if (tier == "FREE")
                    return true;

                // FIX: Thêm tiêu đề "THBIM"
                Interaction.MsgBox(
                    "Your account is pending activation.\n" +
                    "Please open THBIM AutoUpdate app and enter a License Key.",
                    MsgBoxStyle.Information, "THBIM"
                );
                return false;
            }

            // FIX: Thêm tiêu đề "THBIM"
            Interaction.MsgBox("Please sign in using THBIM AutoUpdate app.", MsgBoxStyle.Information, "THBIM");
            return false;
        }

        public static bool EnsureActivated(object payload)
        {
            if (payload == null) return EnsureActivated();

            try
            {
                System.Collections.Generic.Dictionary<string, object> map = null;

                if (payload is string sjson)
                    map = Deserialize(sjson);
                else if (payload is System.Collections.Generic.Dictionary<string, object> dict)
                    map = dict;
                else
                {
                    var json = JsonSerializer.Serialize(payload);
                    map = Deserialize(json);
                }
                if (map == null) map = new System.Collections.Generic.Dictionary<string, object>();

                string tier = map.TryGetValue("tier", out var tr) ? tr?.ToString() : null;
                string expYMD = map.TryGetValue("exp", out var ex) ? ex?.ToString() : null;
                string email = map.TryGetValue("email", out var em) ? em?.ToString() : null;
                string name = map.TryGetValue("fullName", out var fn) ? fn?.ToString() : null;

                if (!TryReadSession(out var ses))
                    ses = new SessionCache();

                if (!string.IsNullOrWhiteSpace(email)) ses.Email = email;
                if (!string.IsNullOrWhiteSpace(name)) ses.FullName = name;
                if (!string.IsNullOrWhiteSpace(tier)) ses.Tier = tier.ToUpperInvariant();
                if (expYMD != null) ses.PremiumExpYMD = expYMD;

                if (!WriteSession(ses)) return false;
                UpdateVerifyStamp();

                Log($"EnsureActivated(payload): {ses.Email}, tier={ses.Tier}, exp={ses.PremiumExpYMD}");
                return true;
            }
            catch (Exception ex)
            {
                Log("EnsureActivated(object) EX: " + ex);
                return false;
            }
        }

        public static string GetCurrentTokenOrNull()
        {
            return TryReadSession(out var ses) ? (ses.Token ?? null) : null;
        }

        public struct LocalLicenseStatus
        {
            public bool HasCache;
            public string Email;
            public string FullName;
            public string Tier;
            public DateTime Exp;
            public bool IsValid;
        }

        public static LocalLicenseStatus GetLocalStatus()
        {
            if (TryReadSession(out var ses))
            {
                var exp = ParseDate(ses.PremiumExpYMD) ?? DateTime.MinValue;
                return new LocalLicenseStatus
                {
                    HasCache = true,
                    Email = ses.Email,
                    FullName = ses.FullName,
                    Tier = string.IsNullOrWhiteSpace(ses.Tier) ? "FREE" : ses.Tier.ToUpperInvariant(),
                    Exp = exp,
                    IsValid = !string.IsNullOrWhiteSpace(ses.Token)
                };
            }

            if (TryReadLegacyCacheInfo(out var email, out var exp2))
                return new LocalLicenseStatus
                {
                    HasCache = true,
                    Email = email,
                    FullName = null,
                    Tier = "FREE",
                    Exp = exp2,
                    IsValid = false
                };

            return new LocalLicenseStatus
            {
                HasCache = false,
                Email = null,
                FullName = null,
                Tier = null,
                Exp = DateTime.MinValue,
                IsValid = false
            };
        }

        public static void Logout()
        {
            try
            {
                var p1 = Path.Combine(AppFolder, SessionFile);
                if (File.Exists(p1)) File.Delete(p1);

                var p2 = Path.Combine(AppFolder, LegacyCacheFile);
                if (File.Exists(p2)) File.Delete(p2);

                var p3 = Path.Combine(AppFolder, VerifyStampFile);
                if (File.Exists(p3)) File.Delete(p3);

                Log("Logout: cleared session and legacy cache.");
            }
            catch (Exception ex) { Log("Logout EX: " + ex); }
        }

        public struct Profile
        {
            public string Email;
            public string FullName;
            public string Tier;
            public string PremiumExpYMD;
        }

        public struct LoginResult
        {
            public bool Ok;
            public string Error;
            public string Token;
            public Profile Profile;
        }

        public static LoginResult TryLoginPwd(string email, string password, string userAgent = null)
        {
            _lastError = null;
            try
            {
                if (string.IsNullOrWhiteSpace(email) || string.IsNullOrWhiteSpace(password))
                    return new LoginResult { Ok = false, Error = "BAD_INPUT" };

                var url = GetServerUrl();
                using (var http = new HttpClient() { Timeout = TimeSpan.FromSeconds(20) })
                {
                    var payload = new
                    {
                        action = "LOGIN_PWD",
                        email = email.Trim(),
                        password = password,
                        userAgent = userAgent ?? Environment.MachineName
                    };
                    var json = JsonSerializer.Serialize(payload);
                    var resp = http.PostAsync(url, new StringContent(json, Encoding.UTF8, "application/json")).Result;
                    var text = resp.Content.ReadAsStringAsync().Result;
                    Log($"LOGIN_PWD {(int)resp.StatusCode} {resp.StatusCode}, body={text}");

                    if (!resp.IsSuccessStatusCode)
                        return new LoginResult { Ok = false, Error = "HTTP_" + (int)resp.StatusCode };

                    var obj = Deserialize(text);
                    if (!(obj.TryGetValue("ok", out var ok) && IsTrue(ok)))
                        return new LoginResult
                        {
                            Ok = false,
                            Error = obj.ContainsKey("error") ? obj["error"]?.ToString() : "server_not_ok"
                        };

                    var token = obj.TryGetValue("token", out var t) ? (t?.ToString() ?? "") : "";
                    var tier = obj.TryGetValue("tier", out var tr) ? (tr?.ToString() ?? "") : "";
                    var exp = obj.TryGetValue("premiumExp", out var ex) ? (ex?.ToString() ?? "") : "";
                    var mail = obj.TryGetValue("email", out var em) ? (em?.ToString() ?? email) : email;
                    var name = obj.TryGetValue("fullName", out var fn) ? (fn?.ToString() ?? "") : "";

                    if (string.IsNullOrWhiteSpace(token))
                        return new LoginResult { Ok = false, Error = "NO_TOKEN" };

                    var ses = new SessionCache
                    {
                        Token = token,
                        Email = mail,
                        FullName = name,
                        Tier = string.IsNullOrWhiteSpace(tier) ? "FREE" : tier,
                        PremiumExpYMD = exp,
                        IssuedAt = DateTime.UtcNow.ToString("o")
                    };
                    if (!WriteSession(ses))
                        return new LoginResult { Ok = false, Error = "WRITE_FAIL" };

                    UpdateVerifyStamp();
                    return new LoginResult
                    {
                        Ok = true,
                        Token = token,
                        Profile = new Profile
                        {
                            Email = mail,
                            FullName = name,
                            Tier = ses.Tier,
                            PremiumExpYMD = ses.PremiumExpYMD
                        }
                    };
                }
            }
            catch (Exception ex)
            {
                _lastError = ex.Message;
                Log("LOGIN_PWD EX: " + ex);
                return new LoginResult { Ok = false, Error = ex.Message };
            }
        }

        public static bool GetProfile(string token, out Profile p)
        {
            p = default;
            try
            {
                var url = GetServerUrl();
                using (var http = new HttpClient() { Timeout = TimeSpan.FromSeconds(20) })
                {
                    var payload = new { action = "GET_PROFILE", token = token };
                    var json = JsonSerializer.Serialize(payload);
                    var resp = http.PostAsync(url, new StringContent(json, Encoding.UTF8, "application/json")).Result;
                    var text = resp.Content.ReadAsStringAsync().Result;
                    Log($"GET_PROFILE {(int)resp.StatusCode} {resp.StatusCode}, body={text}");
                    if (!resp.IsSuccessStatusCode) return false;

                    var obj = Deserialize(text);
                    if (!(obj.TryGetValue("ok", out var ok) && IsTrue(ok))) return false;

                    p.Email = obj.TryGetValue("email", out var em) ? (em?.ToString() ?? "") : "";
                    p.FullName = obj.TryGetValue("fullName", out var fn) ? (fn?.ToString() ?? "") : "";
                    p.Tier = obj.TryGetValue("tier", out var tr) ? (tr?.ToString() ?? "FREE") : "FREE";
                    p.PremiumExpYMD = obj.TryGetValue("premiumExp", out var ex) ? (ex?.ToString() ?? "") : "";
                    return true;
                }
            }
            catch (Exception ex) { Log("GET_PROFILE EX: " + ex); return false; }
        }

        public static bool VerifyAccountByEmail(string email, out string tier, out string expYMD)
        {
            tier = null; expYMD = null;
            try
            {
                var url = GetServerUrl();
                using (var http = new HttpClient() { Timeout = TimeSpan.FromSeconds(20) })
                {
                    var payload = new { action = "VERIFY_ACCOUNT", email = email };
                    var json = JsonSerializer.Serialize(payload);
                    var resp = http.PostAsync(url, new StringContent(json, Encoding.UTF8, "application/json")).Result;
                    var text = resp.Content.ReadAsStringAsync().Result;
                    Log($"VERIFY_ACCOUNT {(int)resp.StatusCode} {resp.StatusCode}, body={text}");
                    if (!resp.IsSuccessStatusCode) return false;

                    var obj = Deserialize(text);
                    if (!(obj.TryGetValue("ok", out var ok) && IsTrue(ok))) return false;

                    tier = obj.TryGetValue("tier", out var tr) ? (tr?.ToString() ?? "FREE") : "FREE";
                    expYMD = obj.TryGetValue("exp", out var ex) ? (ex?.ToString() ?? "") : "";
                    return true;
                }
            }
            catch (Exception ex) { Log("VERIFY_ACCOUNT EX: " + ex); return false; }
        }

        public static string GetServerUrlPublic() => GetServerUrl();
        public static string GetCurrentProductId() => Product;
        public static string GetCurrentMachineId() => MachineIdHelper.Get();
        public static string GetCachedEmailOrNull() => GetLocalStatus().Email;

        private class SessionCache
        {
            public string Token { get; set; }
            public string Email { get; set; }
            public string FullName { get; set; }
            public string Tier { get; set; }
            public string PremiumExpYMD { get; set; }
            public string IssuedAt { get; set; }
        }

        private static bool WriteSession(SessionCache ses)
        {
            try
            {
                Directory.CreateDirectory(AppFolder);
                var json = JsonSerializer.Serialize(ses);
                var enc = ProtectedData.Protect(Encoding.UTF8.GetBytes(json), null, DataProtectionScope.LocalMachine);
                File.WriteAllBytes(Path.Combine(AppFolder, SessionFile), enc);
                Log($"Session write OK: {ses.Email}, tier={ses.Tier}, exp={ses.PremiumExpYMD}");
                return true;
            }
            catch (Exception ex) { Log("WRITE SESSION EX: " + ex); return false; }
        }

        private static bool TryReadSession(out SessionCache ses)
        {
            ses = null;
            try
            {
                var p = Path.Combine(AppFolder, SessionFile);
                if (!File.Exists(p)) return false;
                var enc = File.ReadAllBytes(p);
                var raw = ProtectedData.Unprotect(enc, null, DataProtectionScope.LocalMachine);
                var txt = Encoding.UTF8.GetString(raw);
                ses = JsonSerializer.Deserialize<SessionCache>(txt);
                return ses != null && !string.IsNullOrWhiteSpace(ses.Email);
            }
            catch (Exception ex) { Log("READ SESSION EX: " + ex); return false; }
        }

        private static bool TryReadLegacyCacheInfo(out string email, out DateTime exp)
        {
            email = null; exp = DateTime.MinValue;
            var path = Path.Combine(AppFolder, LegacyCacheFile);
            if (!File.Exists(path)) return false;

            try
            {
                var enc = File.ReadAllBytes(path);
                var raw = ProtectedData.Unprotect(enc, null, DataProtectionScope.LocalMachine);
                var txt = Encoding.UTF8.GetString(raw).Trim();

                try
                {
                    var obj = JsonSerializer.Deserialize<System.Collections.Generic.Dictionary<string, object>>(txt);
                    if (obj != null)
                    {
                        if (obj.TryGetValue("email", out var e) && e != null) email = e.ToString();
                        if (obj.TryGetValue("exp", out var x) && x != null)
                        {
                            var d = ParseDate(x.ToString());
                            if (d != null) { exp = d.Value.Date; return true; }
                        }
                    }
                }
                catch { /* fallback */ }

                if (DateTime.TryParse(txt, out var d2))
                {
                    email = "(legacy)";
                    exp = d2.Date;
                    return true;
                }
            }
            catch (Exception ex) { Log("READ LEGACY EX: " + ex); }

            return false;
        }

        private static bool NeedVerify()
        {
            try
            {
                if (VerifyIntervalDays <= 0) return true;
                var p = Path.Combine(AppFolder, VerifyStampFile);
                if (!File.Exists(p)) return true;
                var s = File.ReadAllText(p).Trim();
                if (!DateTime.TryParseExact(s, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var last))
                {
                    if (!DateTime.TryParse(s, out last)) return true;
                }
                return DateTime.UtcNow.Date >= last.Date.AddDays(VerifyIntervalDays);
            }
            catch { return true; }
        }

        private static void UpdateVerifyStamp()
        {
            try
            {
                Directory.CreateDirectory(AppFolder);
                var p = Path.Combine(AppFolder, VerifyStampFile);
                File.WriteAllText(p, DateTime.UtcNow.ToString("yyyy-MM-dd"));
            }
            catch { }
        }

        // FIX: Helper cũng thêm tiêu đề "THBIM"
        private static void MsgInfo(string msg) => Interaction.MsgBox(msg, MsgBoxStyle.Information, "THBIM");

        private static DateTime? ParseDate(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return null;
            s = s.Trim();
            if (DateTime.TryParseExact(s, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var d1))
                return d1.Date;

            var idx = s.IndexOf(" (", StringComparison.Ordinal);
            if (idx > 0) s = s.Substring(0, idx);
            s = s.Replace("GMT", "").Trim();
            if (DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var d2))
                return d2.Date;

            return null;
        }

        public static bool IsLoggedIn()
        {
            var s = GetLocalStatus();
            return s.IsValid;
        }

        public static bool HasPremium()
        {
            var s = GetLocalStatus();
            if (!s.IsValid) return false;
            var isPremium = string.Equals(s.Tier ?? "FREE", "PREMIUM", StringComparison.OrdinalIgnoreCase);
            if (!isPremium) return false;
            if (s.Exp != DateTime.MinValue && DateTime.UtcNow.Date > s.Exp.Date) return false;
            return true;
        }

        public static bool EnsurePremium(string customMsg = null)
        {
            if (!IsLoggedIn())
            {
                // FIX: Thêm tiêu đề "THBIM"
                Interaction.MsgBox("Please sign in using THBIM AutoUpdate app.", MsgBoxStyle.Information, "THBIM");
                return false;
            }
            if (!HasPremium())
            {
                var msg = customMsg ??
                          "This is a PREMIUM feature.\nPlease open THBIM AutoUpdate app and enter a Premium key.";
                // FIX: Thêm tiêu đề "THBIM"
                Interaction.MsgBox(msg, MsgBoxStyle.Information, "THBIM");
                return false;
            }
            return true;
        }

        private static bool IsTrue(object v)
        {
            if (v == null) return false;
            var s = v.ToString().Trim().ToLowerInvariant();
            return s == "true" || s == "1" || s == "ok" || s == "yes";
        }

        private static System.Collections.Generic.Dictionary<string, object> Deserialize(string json)
        {
            try
            {
                return JsonSerializer.Deserialize<System.Collections.Generic.Dictionary<string, object>>(json);
            }
            catch (Exception ex)
            {
                Log("DESERIALIZE EX: " + ex);
                return new System.Collections.Generic.Dictionary<string, object>();
            }
        }

        private static string GetServerUrl()
        {
            var url = (LicenseServer ?? "").Trim();
            if (!Uri.TryCreate(url, UriKind.Absolute, out var u) ||
                (u.Scheme != Uri.UriSchemeHttps && u.Scheme != Uri.UriSchemeHttp))
            {
                Log("Invalid LicenseServer URL: '" + url + "'");
                // FIX: Thêm tiêu đề "THBIM"
                Interaction.MsgBox("Invalid license server URL:\n" + url, MsgBoxStyle.Critical, "THBIM");
                throw new InvalidOperationException("Invalid LicenseServer URL");
            }
            return u.ToString();
        }

        private static void Log(string msg)
        {
            try
            {
                Directory.CreateDirectory(AppFolder);
                var p = Path.Combine(AppFolder, "licensing.log");
                File.AppendAllText(p,
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {msg}\r\n",
                    Encoding.UTF8);
            }
            catch
            {
                try
                {
                    var fallback = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                        "THBIM", "Licensing");
                    Directory.CreateDirectory(fallback);
                    var p2 = Path.Combine(fallback, "licensing.log");
                    File.AppendAllText(p2,
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {msg}\r\n",
                        Encoding.UTF8);
                }
                catch { }
            }
        }
    }
}