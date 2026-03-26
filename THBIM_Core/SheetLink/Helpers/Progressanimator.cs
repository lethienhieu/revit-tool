using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;

namespace THBIM.Helpers
{
    /// <summary>
    /// Điều phối animation progress bar trong SheetLinkWindow.
    /// Chạy trên UI thread qua Dispatcher — an toàn với Revit API.
    /// </summary>
    public class ProgressAnimator
    {
        private readonly Border _fill;
        private readonly TextBlock _label;

        // Màu sắc theo trạng thái
        private static readonly SolidColorBrush BrushIdle = new(Color.FromRgb(0xD8, 0xF0, 0xD8));
        private static readonly SolidColorBrush BrushActive = new(Color.FromRgb(0xC8, 0xEE, 0xC8));
        private static readonly SolidColorBrush BrushDone = new(Color.FromRgb(0xA8, 0xE0, 0xA8));
        private static readonly SolidColorBrush BrushError = new(Color.FromRgb(0xFD, 0xEA, 0xEA));
        private static readonly SolidColorBrush BrushWarning = new(Color.FromRgb(0xFD, 0xF0, 0xCC));

        private static readonly SolidColorBrush FgIdle = new(Color.FromRgb(0x2A, 0x6E, 0x2A));
        private static readonly SolidColorBrush FgError = new(Color.FromRgb(0x92, 0x22, 0x22));
        private static readonly SolidColorBrush FgWarning = new(Color.FromRgb(0x8A, 0x60, 0x10));

        private CancellationTokenSource _cts;

        public ProgressAnimator(Border fillBorder, TextBlock labelBlock)
        {
            _fill = fillBorder ?? throw new ArgumentNullException(nameof(fillBorder));
            _label = labelBlock ?? throw new ArgumentNullException(nameof(labelBlock));
        }

        // ══════════════════════════════════════════════════════════════════
        // PUBLIC API
        // ══════════════════════════════════════════════════════════════════

        /// <summary>Cập nhật progress bar — gọi được từ bất kỳ thread nào.</summary>
        public void SetProgress(int pct, string message = null)
        {
            Application.Current?.Dispatcher.Invoke(() =>
            {
                var msg = message ?? $"Completed   {pct}%";
                _label.Text = msg;

                if (pct >= 100)
                {
                    _fill.Background = BrushDone;
                    _label.Foreground = FgIdle;
                    AnimateFade(_fill, from: 1.0, to: 0.85, durationMs: 400);
                }
                else if (pct > 0)
                {
                    _fill.Background = BrushActive;
                    _label.Foreground = FgIdle;
                }
                else
                {
                    _fill.Background = BrushIdle;
                    _label.Foreground = FgIdle;
                }
            });
        }

        /// <summary>Trạng thái lỗi — đỏ nhạt.</summary>
        public void SetError(string message)
        {
            Application.Current?.Dispatcher.Invoke(() =>
            {
                _fill.Background = BrushError;
                _label.Text = message;
                _label.Foreground = FgError;
                ShakeAnimation(_fill);
            });
        }

        /// <summary>Trạng thái cảnh báo — vàng nhạt.</summary>
        public void SetWarning(string message)
        {
            Application.Current?.Dispatcher.Invoke(() =>
            {
                _fill.Background = BrushWarning;
                _label.Text = message;
                _label.Foreground = FgWarning;
            });
        }

        /// <summary>Bắt đầu animation "đang xử lý" (pulse) cho đến khi gọi SetProgress/SetError.</summary>
        public void StartIndeterminate(string message = "Processing...")
        {
            _cts?.Cancel();
            _cts = new CancellationTokenSource();
            var token = _cts.Token;

            Application.Current?.Dispatcher.Invoke(() =>
            {
                _fill.Background = BrushActive;
                _label.Text = message;
                _label.Foreground = FgIdle;
            });

            Task.Run(async () =>
            {
                int dot = 0;
                while (!token.IsCancellationRequested)
                {
                    await Task.Delay(500, token).ContinueWith(_ => { });
                    if (token.IsCancellationRequested) break;

                    var dots = new string('.', (dot % 3) + 1).PadRight(3);
                    Application.Current?.Dispatcher.Invoke(() =>
                        _label.Text = $"{message}{dots}");
                    dot++;
                }
            }, token);
        }

        /// <summary>Dừng animation indeterminate.</summary>
        public void StopIndeterminate()
        {
            _cts?.Cancel();
            _cts = null;
        }

        /// <summary>Reset về trạng thái ban đầu.</summary>
        public void Reset()
        {
            StopIndeterminate();
            Application.Current?.Dispatcher.Invoke(() =>
            {
                _fill.Background = BrushIdle;
                _label.Text = "Completed   0%";
                _label.Foreground = FgIdle;
                _fill.Opacity = 1.0;
            });
        }

        // ══════════════════════════════════════════════════════════════════
        // PRIVATE ANIMATIONS
        // ══════════════════════════════════════════════════════════════════

        private static void AnimateFade(UIElement el, double from, double to, int durationMs)
        {
            var anim = new DoubleAnimation(from, to,
                new Duration(TimeSpan.FromMilliseconds(durationMs)))
            {
                EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut }
            };
            el.BeginAnimation(UIElement.OpacityProperty, anim);
        }

        private static void ShakeAnimation(UIElement el)
        {
            var transform = new TranslateTransform();
            el.RenderTransform = transform;

            var anim = new DoubleAnimationUsingKeyFrames();
            anim.KeyFrames.Add(new EasingDoubleKeyFrame(0, KeyTime.FromTimeSpan(TimeSpan.FromMilliseconds(0))));
            anim.KeyFrames.Add(new EasingDoubleKeyFrame(-4, KeyTime.FromTimeSpan(TimeSpan.FromMilliseconds(80))));
            anim.KeyFrames.Add(new EasingDoubleKeyFrame(4, KeyTime.FromTimeSpan(TimeSpan.FromMilliseconds(160))));
            anim.KeyFrames.Add(new EasingDoubleKeyFrame(-3, KeyTime.FromTimeSpan(TimeSpan.FromMilliseconds(240))));
            anim.KeyFrames.Add(new EasingDoubleKeyFrame(3, KeyTime.FromTimeSpan(TimeSpan.FromMilliseconds(320))));
            anim.KeyFrames.Add(new EasingDoubleKeyFrame(0, KeyTime.FromTimeSpan(TimeSpan.FromMilliseconds(400))));

            transform.BeginAnimation(TranslateTransform.XProperty, anim);
        }
    }
}