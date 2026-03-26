using System;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media.Animation;
using System.Windows.Navigation;
using System.Windows.Threading;
using THBIM.Licensing;

namespace THBIM.Tools
{
    public partial class LoginWindow : Window
    {
        private DispatcherTimer _timer;
        private double _prog; private DateTime _last;
        private const double CAP = 0.94, SPEED = 0.30;

        public LoginWindow()
        {
            InitializeComponent();
            Topmost = true;
            Prefill();
        }

        private void Prefill()
        {
            var s = LicenseManager.GetLocalStatus();
            if (!string.IsNullOrWhiteSpace(s.Email)) EmailBox.Text = s.Email;
        }

        // Busy overlay
        private void SetBusy(bool on, string title = null, string note = null)
        {
            BusyOverlay.Visibility = on ? Visibility.Visible : Visibility.Collapsed;
            this.IsEnabled = !on;
            LoginBtn.IsEnabled = !on; SignupBtn.IsEnabled = !on;
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

        private async void Login_Click(object sender, RoutedEventArgs e)
        {
            var email = (EmailBox.Text ?? "").Trim();
            var pwd = (PwdBox.Password ?? "").Trim();
            if (string.IsNullOrWhiteSpace(email) || string.IsNullOrWhiteSpace(pwd))
            {
                MessageBox.Show("Vui lòng nhập email và mật khẩu.", "THBIM",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                SetBusy(true, "Đang đăng nhập…");
                var result = await Task.Run(() => LicenseManager.TryLoginPwd(email, pwd, Environment.MachineName));
                FinishSmooth();

                if (!result.Ok)
                {
                    string err = (result.Error ?? "").ToUpperInvariant();
                    string msg =
                        err == "EMAIL_NOT_FOUND" ? "Email không tồn tại." :
                        err == "WRONG_PASSWORD" ? "Mật khẩu không đúng." :
                        err == "BAD_INPUT" ? "Thiếu email/mật khẩu." :
                        err.StartsWith("HTTP_") ? "Không kết nối được máy chủ (HTTP)." :
                        string.IsNullOrWhiteSpace(result.Error) ? "Đăng nhập thất bại." : result.Error;
                    MessageBox.Show(msg, "THBIM", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // OK
                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                FinishSmooth();
                MessageBox.Show("Lỗi khi đăng nhập:\n" + ex.Message, "THBIM",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                await Task.Delay(200);
                SetBusy(false);
            }
        }
        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = e.Uri.AbsoluteUri,
                UseShellExecute = true
            });
            e.Handled = true;
        }

        private void Signup_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "https://thbim.pages.dev",
                    UseShellExecute = true
                });
            }
            catch { }
        }

        private void Root_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter && LoginBtn.IsEnabled) Login_Click(sender, e);
        }
    }
}
