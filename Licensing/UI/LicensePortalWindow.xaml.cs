using System;
using System.Collections.Generic; // Added to use Dictionary
using System.Net.Http;
using System.Text;
using System.Text.Json; // Replaces System.Web.Script.Serialization
using System.Windows;
using System.Windows.Media.Animation;
using System.Windows.Threading;
using THBIM.Licensing;

namespace THBIM.Tools
{
    public partial class LicensePortalWindow : Window
    {
        private const string ACTION_APPLY_KEY = "APPLY_KEY"; // GAS returns { ok, tier, exp, message }
        private const string ACTION_VERIFY = "VERIFY";

        private DispatcherTimer _timer;
        private double _prog; private DateTime _last;
        private const double CAP = 0.94, SPEED = 0.30;

        public LicensePortalWindow()
        {
            InitializeComponent();
            LicenseManager.SetProduct("TH-SUITE");
            LoadProfile();
        }
        private static string MapTierToDisplay(string tier)
        {
            if (string.IsNullOrWhiteSpace(tier)) return "Pending Activation";
            var t = tier.Trim().ToUpperInvariant();
            if (t == "PENDING") return "Pending Activation";
            if (t == "FREE") return "FREE";
            if (t == "PREMIUM") return "PREMIUM";
            return t; // fallback
        }

        private void LoadProfile()
        {
            var s = LicenseManager.GetLocalStatus();
            FullNameText.Text = string.IsNullOrWhiteSpace(s.FullName) ? "Full Name" : s.FullName;
            EmailText.Text = string.IsNullOrWhiteSpace(s.Email) ? "user@mail.com" : s.Email;
            MachineText.Text = LicenseManager.GetCurrentMachineId();

            // DEFAULT: Pending Activation if no tier
            var tierRaw = string.IsNullOrWhiteSpace(s.Tier) ? "PENDING" : s.Tier;
            StatusPill.Text = " " + MapTierToDisplay(tierRaw) + " ";

            ExpireText.Text = FormatExpText(s.Exp);

        }

        /* ===== Busy overlay ===== */
        private void SetBusy(bool on, string title = null, string note = null)
        {
            BusyOverlay.Visibility = on ? Visibility.Visible : Visibility.Collapsed;
            this.IsEnabled = !on;
            if (title != null) BusyTitle.Text = title;
            if (note != null) BusyNote.Text = note;
            if (on) { _prog = 0; SetScale(0); StartTimer(); } else StopTimer();
        }
        private void StartTimer()
        {
            _last = DateTime.UtcNow;
            if (_timer == null) { _timer = new DispatcherTimer(DispatcherPriority.Render) { Interval = TimeSpan.FromMilliseconds(16) }; }
            _timer.Tick -= OnTick; _timer.Tick += OnTick; _timer.Start();
        }
        private void StopTimer() { if (_timer != null) { _timer.Tick -= OnTick; _timer.Stop(); } }
        private void OnTick(object s, EventArgs e) { var now = DateTime.UtcNow; var dt = (now - _last).TotalSeconds; _last = now; _prog = Math.Min(CAP, _prog + SPEED * dt); SetScale(_prog); }
        private void SetScale(double v) { BusyScale.ScaleX = Math.Max(0, Math.Min(1, v)); }
        private void FinishSmooth()
        {
            StopTimer();
            BusyScale.BeginAnimation(
                System.Windows.Media.ScaleTransform.ScaleXProperty,
                new DoubleAnimation
                {
                    From = BusyScale.ScaleX,
                    To = 1.0,
                    Duration = TimeSpan.FromMilliseconds(350),
                    EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut }
                },
                HandoffBehavior.SnapshotAndReplace);
        }

        /* ===== Buttons ===== */
        private void Feedback_Click(object sender, RoutedEventArgs e)
        {
            var win = new FeedbackWindow();
            new System.Windows.Interop.WindowInteropHelper(win)
                .Owner = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
            win.ShowDialog();
        }

        private void Logout_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Logging out will clear the login session on this machine. Continue?",
                                "THBIM", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) return;
            LicenseManager.Logout();
            MessageBox.Show("Logged out. Please log in again to use the service.", "THBIM",
                            MessageBoxButton.OK, MessageBoxImage.Information);
            Close();
        }

        private async void Activate_Click(object sender, RoutedEventArgs e)
        {
            var key = (KeyBox.Text ?? "").Trim();
            if (string.IsNullOrWhiteSpace(key))
            {
                MessageBox.Show("Please enter the key.", "THBIM", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var token = LicenseManager.GetCurrentTokenOrNull();
            if (string.IsNullOrWhiteSpace(token))
            {
                MessageBox.Show("Login session expired. Please log in again.", "THBIM",
                                MessageBoxButton.OK, MessageBoxImage.Warning);
                Close(); return;
            }

            try
            {
                ActivateBtn.IsEnabled = false;
                SetBusy(true, "Activating…");

                using (var http = new HttpClient() { Timeout = TimeSpan.FromSeconds(20) })
                {
                    var payload = new
                    {
                        action = ACTION_APPLY_KEY,
                        token = token,
                        key = key,
                        machineId = LicenseManager.GetCurrentMachineId()
                    };
                    // FIX: Use System.Text.Json
                    var json = JsonSerializer.Serialize(payload);
                    var resp = await http.PostAsync(LicenseManager.GetServerUrlPublic(),
                                                    new StringContent(json, Encoding.UTF8, "application/json"));
                    var text = await resp.Content.ReadAsStringAsync();
                    FinishSmooth();

                    // FIX: Deserialize to Dictionary instead of dynamic for .NET 8 compatibility
                    var obj = JsonSerializer.Deserialize<Dictionary<string, object>>(text);

                    // Logic check OK remains same, adjusted how value is retrieved from Dictionary
                    bool ok = false;
                    if (obj != null && obj.ContainsKey("ok"))
                    {
                        var okVal = obj["ok"]?.ToString().ToLower();
                        ok = (okVal == "true" || okVal == "1");
                    }

                    if (!ok)
                    {
                        var err = (obj != null && obj.ContainsKey("error")) ? (obj["error"]?.ToString() ?? "") : "";
                        MessageBox.Show("Activation failed.\n" + err, "THBIM", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    string tier = (obj != null && obj.ContainsKey("tier")) ? (obj["tier"]?.ToString() ?? "FREE") : "FREE";
                    string exp = (obj != null && obj.ContainsKey("exp")) ? (obj["exp"]?.ToString() ?? "") : "";
                    string msg = (obj != null && obj.ContainsKey("message")) ? (obj["message"]?.ToString() ?? "") : "";

                    // DISPLAY BY MAP
                    StatusPill.Text = " " + MapTierToDisplay(tier) + " ";
                    ExpireText.Text = FormatExpText(exp);
                    ResultMsg.Visibility = Visibility.Visible;
                    ResultMsg.Text = string.IsNullOrWhiteSpace(msg)
                        ? (tier.ToUpper() == "PREMIUM"
                            ? $"✓ Success: PREMIUM until {ExpireText.Text}"
                            : (tier.ToUpper() == "FREE"
                                ? "✓ Success: FREE activated"
                                : "✓ Status updated"))
                        : msg;

                    // >>> REFRESH PROFILE FROM SERVER
                    try
                    {
                        var email = LicenseManager.GetCachedEmailOrNull();
                        if (!string.IsNullOrWhiteSpace(email))
                        {
                            var verifyPayload = new { action = "VERIFY_ACCOUNT", email = email };
                            // FIX: Use System.Text.Json
                            var verifyJson = JsonSerializer.Serialize(verifyPayload);

                            var verifyResp = await http.PostAsync(
                                LicenseManager.GetServerUrlPublic(),
                                new StringContent(verifyJson, Encoding.UTF8, "application/json")
                            );
                            var verifyText = await verifyResp.Content.ReadAsStringAsync();   // ← Original JSON from GAS

                            // WRITE CACHE using ORIGINAL JSON
                            LicenseManager.EnsureActivated(verifyText);

                            // READ BACK FROM CACHE just written → set UI
                            var s = LicenseManager.GetLocalStatus();
                            StatusPill.Text = " " + MapTierToDisplay(s.Tier) + " ";
                            ExpireText.Text = FormatExpText(s.Exp);
                        }
                    }
                    catch
                    {
                        // silent; UI reported success above
                    }
                }
            }
            catch (Exception ex)
            {
                FinishSmooth();
                MessageBox.Show("Activation error:\n" + ex.Message, "THBIM",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                await System.Threading.Tasks.Task.Delay(200);
                SetBusy(false);
                ActivateBtn.IsEnabled = true;
            }
        }

        private const string LIFETIME_YMD = "9999-12-31";
        private static string FormatExpText(DateTime exp)
            => (exp == DateTime.MinValue) ? "-" : (exp.Year >= 9999 ? "LIFETIME" : exp.ToString("yyyy-MM-dd"));
        private static string FormatExpText(string expYmd)
            => string.IsNullOrWhiteSpace(expYmd) ? "-" : (expYmd.Trim() == LIFETIME_YMD ? "LIFETIME" : expYmd.Trim());
    }
}