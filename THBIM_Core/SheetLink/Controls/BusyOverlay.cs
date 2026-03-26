using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;

namespace THBIM.Controls
{
    public class BusyOverlay : Grid
    {
        private readonly TextBlock _msg;
        private readonly Ellipse   _spinner;

        public BusyOverlay()
        {
            Visibility = Visibility.Collapsed;
            Background = new SolidColorBrush(Color.FromArgb(0xAA, 0x00, 0x00, 0x00));
            SetZIndex(this, 9999);

            var panel = new Border
            {
                Background      = new SolidColorBrush(Color.FromRgb(0xF8, 0xF8, 0xF8)),
                BorderBrush     = new SolidColorBrush(Color.FromRgb(0xD0, 0xD0, 0xD0)),
                BorderThickness = new Thickness(1),
                CornerRadius    = new CornerRadius(8),
                Padding         = new Thickness(32, 24, 32, 24),
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment   = VerticalAlignment.Center,
                MinWidth        = 220
            };

            var inner = new StackPanel
            {
                Orientation         = Orientation.Vertical,
                HorizontalAlignment = HorizontalAlignment.Center
            };

            _spinner = new Ellipse
            {
                Width  = 36, Height = 36,
                Stroke = new SolidColorBrush(Color.FromRgb(0x2A, 0x6E, 0x2A)),
                StrokeThickness = 3,
                StrokeDashArray = new DoubleCollection { 8, 4 },
                HorizontalAlignment   = HorizontalAlignment.Center,
                Margin                = new Thickness(0, 0, 0, 14),
                RenderTransformOrigin = new Point(0.5, 0.5),
                RenderTransform       = new RotateTransform()
            };

            ((RotateTransform)_spinner.RenderTransform)
                .BeginAnimation(RotateTransform.AngleProperty,
                    new DoubleAnimation(0, 360,
                        new Duration(TimeSpan.FromSeconds(1)))
                    { RepeatBehavior = RepeatBehavior.Forever });

            _msg = new TextBlock
            {
                Text                = "Processing...",
                FontSize            = 13,
                FontFamily          = new FontFamily("Segoe UI"),
                Foreground          = new SolidColorBrush(Color.FromRgb(0x1A, 0x1A, 0x1A)),
                TextAlignment       = TextAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                TextWrapping        = TextWrapping.Wrap,
                MaxWidth            = 200
            };

            inner.Children.Add(_spinner);
            inner.Children.Add(_msg);
            panel.Child = inner;
            Children.Add(panel);
        }

        public void Show(string message = "Processing...")
        {
            _msg.Text  = message;
            Visibility = Visibility.Visible;
            BeginAnimation(OpacityProperty,
                new DoubleAnimation(0, 1, new Duration(TimeSpan.FromMilliseconds(150))));
        }

        public void UpdateMessage(string message)
        {
            Application.Current?.Dispatcher.Invoke(() => _msg.Text = message);
        }

        public void Hide()
        {
            Application.Current?.Dispatcher.Invoke(() =>
            {
                var a = new DoubleAnimation(1, 0,
                    new Duration(TimeSpan.FromMilliseconds(150)));
                a.Completed += (_, _) => Visibility = Visibility.Collapsed;
                BeginAnimation(OpacityProperty, a);
            });
        }
    }
}
