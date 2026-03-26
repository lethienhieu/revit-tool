using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Reflection; // Cần thêm namespace này để dùng Assembly & GetName
using System.Text;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Animation;
using System.Windows.Threading;
using THBIM.Licensing;

namespace THBIM.Tools
{
    internal class ToolChoice
    {
        public string Id { get; set; }        // "ALL" hoặc <tên DLL không .dll>
        public string Display { get; set; }   // text hiển thị
        public string DllPath { get; set; }   // null nếu ALL
        public string Version { get; set; }   // version gửi kèm
    }

    public partial class FeedbackWindow : Window
    {
        private static readonly string[] ExcludePrefixes = new[]
        {
            "autodesk.", "microsoft.", "system.", "newtonsoft.", "epplus", "google.", "grpc.", "protobuf",
            "presentationframework", "presentationcore", "windowsbase", "reachframework",
            "revitapi", "revitapimacros", "adwindows", "icsharpcode."
        };

        private void ApplyLoginState()
        {
            var st = LicenseManager.GetLocalStatus();
            EmailBox.Text = string.IsNullOrWhiteSpace(st.Email) ? "(chưa đăng nhập)" : st.Email;
            bool canSend = !string.IsNullOrWhiteSpace(st.Email) && st.Email.Contains("@");
            SendBtn.IsEnabled = canSend;
        }

        public FeedbackWindow()
        {
            InitializeComponent();
            SubjectBox.Text = "Góp ý người dùng";
            EmailBox.IsReadOnly = true; EmailBox.Focusable = false; EmailBox.IsTabStop = false;
            SubjectBox.IsReadOnly = true; SubjectBox.Focusable = false; SubjectBox.IsTabStop = false;

            ApplyLoginState();
            LoadToolList();
        }

        private void LoadToolList()
        {
            var list = BuildToolChoices();
            ToolCombo.ItemsSource = list;
            ToolCombo.SelectedValuePath = "Id";
            ToolCombo.SelectedIndex = 0; // mặc định ALL - TH SUITE
        }

        // --- ĐÃ SỬA LOGIC: Thay vì quét thư mục, quét AppDomain để tìm tool có tham chiếu Licensing ---
        private List<ToolChoice> BuildToolChoices()
        {
            var list = new List<ToolChoice>();

            // 1. Lấy thông tin của chính file Licensing (làm mốc)
            var licAsm = typeof(LicenseManager).Assembly;
            var licName = licAsm.GetName().Name; // Tên assembly: ví dụ "THBIM.Licensing"
            var licPath = licAsm.Location;
            var licVer = GetProductOrFileVersionSafe(licPath) ?? (licAsm.GetName().Version?.ToString() ?? "1.0.0.0");

            // Thêm lựa chọn "ALL"
            list.Add(new ToolChoice
            {
                Id = "ALL",
                Display = "ALL - TH SUITE",
                DllPath = null,
                Version = licVer
            });

            // 2. Quét tất cả Assembly ĐÃ LOAD trong Revit (AppDomain)
            // Chỉ lấy những cái nào có Reference tới "THBIM.Licensing"
            var loadedAssemblies = AppDomain.CurrentDomain.GetAssemblies();

            foreach (var asm in loadedAssemblies)
            {
                try
                {
                    // Bỏ qua dynamic assembly hoặc assembly không có đường dẫn file thực
                    if (asm.IsDynamic || string.IsNullOrEmpty(asm.Location)) continue;

                    // Bỏ qua chính file Licensing (đã có ở mục ALL rồi)
                    if (asm.GetName().Name == licName) continue;

                    // KIỂM TRA QUAN TRỌNG: Tool này có dùng (tham chiếu) LicenseManager không?
                    // Nếu có -> Nó là tool của bạn. Nếu không -> Nó là rác của Revit/System.
                    var refs = asm.GetReferencedAssemblies();
                    if (refs.Any(r => r.Name == licName))
                    {
                        var path = asm.Location;
                        var fileName = Path.GetFileName(path);
                        var id = Path.GetFileNameWithoutExtension(fileName);

                        // Lấy version file thực tế
                        var ver = GetProductOrFileVersionSafe(path) ?? asm.GetName().Version?.ToString() ?? "1.0.0.0";

                        // Tránh trùng lặp (nếu Revit lỡ load 2 phiên bản)
                        if (list.Any(x => x.Id == id)) continue;

                        list.Add(new ToolChoice
                        {
                            Id = id,
                            Display = $"{id} — v{ver}",
                            DllPath = path,
                            Version = ver
                        });
                    }
                }
                catch
                {
                    // Bỏ qua các file hệ thống bị lỗi truy cập (nếu có)
                }
            }

            // Sắp xếp: ALL đứng đầu, còn lại sort theo tên
            var all = list.First();
            var rest = list.Skip(1).OrderBy(x => x.Id, StringComparer.OrdinalIgnoreCase).ToList();
            var result = new List<ToolChoice> { all };
            result.AddRange(rest);
            return result;
        }

        private static string GetProductOrFileVersionSafe(string path)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return null;
                var fvi = FileVersionInfo.GetVersionInfo(path);
                return string.IsNullOrWhiteSpace(fvi.ProductVersion) ? fvi.FileVersion : fvi.ProductVersion;
            }
            catch { return null; }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e) => Close();

        // ================= Busy overlay =================
        private DispatcherTimer _progressTimer;
        private double _progress;
        private const double _capBeforeDone = 0.94;
        private const double _speedPerSecond = 0.30;
        private DateTime _lastTick;

        private void SetBusy(bool on, string note = null)
        {
            try
            {
                if (BusyOverlay != null) BusyOverlay.Visibility = on ? Visibility.Visible : Visibility.Collapsed;
                if (!string.IsNullOrWhiteSpace(note) && BusyNote != null) BusyNote.Text = note;
                this.IsEnabled = !on;

                if (on) { _progress = 0; SetBarScale(0); StartProgressTimer(); }
                else { StopProgressTimer(); }
            }
            catch { }
        }

        private void StartProgressTimer()
        {
            _lastTick = DateTime.UtcNow;
            if (_progressTimer == null)
            {
                _progressTimer = new DispatcherTimer(DispatcherPriority.Render);
                _progressTimer.Interval = TimeSpan.FromMilliseconds(16);
            }
            _progressTimer.Tick -= OnProgressTick;
            _progressTimer.Tick += OnProgressTick;
            _progressTimer.Start();
        }

        private void StopProgressTimer()
        {
            if (_progressTimer != null) { _progressTimer.Tick -= OnProgressTick; _progressTimer.Stop(); }
        }

        private void OnProgressTick(object sender, EventArgs e)
        {
            var now = DateTime.UtcNow;
            var dt = (now - _lastTick).TotalSeconds;
            _lastTick = now;
            _progress = Math.Min(_capBeforeDone, _progress + _speedPerSecond * dt);
            SetBarScale(_progress);
        }

        private void SetBarScale(double value01)
        {
            try { BusyScale.ScaleX = Math.Max(0, Math.Min(1, value01)); } catch { }
        }

        private void FinishProgressSmooth()
        {
            try
            {
                StopProgressTimer();
                double from = BusyScale?.ScaleX ?? _progress;
                var anim = new DoubleAnimation
                {
                    From = from,
                    To = 1.0,
                    Duration = TimeSpan.FromMilliseconds(350),
                    EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut }
                };
                BusyScale?.BeginAnimation(System.Windows.Media.ScaleTransform.ScaleXProperty, anim, HandoffBehavior.SnapshotAndReplace);
            }
            catch { }
        }

        // ========================= SEND =========================
        private async void Send_Click(object sender, RoutedEventArgs e)
        {
            var email = (EmailBox.Text ?? "").Trim();
            var subject = (SubjectBox.Text ?? "").Trim();
            var message = (BodyBox.Text ?? "").Trim();
            if (!SendBtn.IsEnabled) return;

            if (string.IsNullOrWhiteSpace(message))
            {
                MessageBox.Show("Please enter your feedback.", "THBIM", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                SetBusy(true, "Sending feedback…");
                var choice = ToolCombo.SelectedItem as ToolChoice ?? BuildToolChoices().First();
                var url = LicenseManager.GetServerUrlPublic();
                var machineId = LicenseManager.GetCurrentMachineId();
                var product = choice.Id == "ALL" ? "TH-SUITE" : choice.Id;
                var appVer = choice.Version;
                var host = Environment.MachineName;

                using (var http = new HttpClient() { Timeout = TimeSpan.FromSeconds(20) })
                {
                    var payload = new
                    {
                        action = "REPORT",
                        email = email,
                        product = product,
                        machineId = machineId,
                        subject = string.IsNullOrWhiteSpace(subject) ? $"Feedback: {choice.Id}" : subject,
                        message = message,
                        appVersion = appVer,
                        host = host,
                        extra = choice.Id == "ALL" ? "ALL_TOOLS" : ""
                    };
                    var json = JsonSerializer.Serialize(payload);

                    var resp = await http.PostAsync(url, new StringContent(json, Encoding.UTF8, "application/json"));
                    var text = await resp.Content.ReadAsStringAsync();

                    FinishProgressSmooth();

                    if (!resp.IsSuccessStatusCode || text.IndexOf("\"ok\":true", StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        MessageBox.Show("Fail.\n" + text, "THBIM", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }
                }

                MessageBox.Show($"Sent feedback ({choice.Display}). Thanks!", "THBIM", MessageBoxButton.OK, MessageBoxImage.Information);
                Close();
            }
            catch (Exception ex)
            {
                FinishProgressSmooth();
                MessageBox.Show("Fail send feedback:\n" + ex.Message, "THBIM", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                await System.Threading.Tasks.Task.Delay(200);
                SetBusy(false);
            }
        }
    }
}