using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;

namespace THBIM
{
    public partial class LeaderAngleSettingWindow : Window
    {
        private double _angle;
        private bool _suppressTextChanged;
        private readonly Button[] _presetButtons;

        public LeaderAngleSettingWindow()
        {
            InitializeComponent();
            _presetButtons = new[] { btn30, btn45, btn60, btn90 };
            _angle = LeaderAngleSettings.AngleDegrees;
            UpdateUI();
        }

        private void Preset_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn)
            {
                var text = btn.Content.ToString().Replace("°", "");
                if (double.TryParse(text, out double val))
                {
                    _angle = val;
                    UpdateUI();
                }
            }
        }

        private void TxtCustom_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_suppressTextChanged) return;
            if (double.TryParse(txtCustom.Text, out double val) && val >= 0 && val <= 90)
            {
                _angle = val;
                UpdatePresetHighlight();
                DrawPreview();
                UpdateLabel();
            }
        }

        private void Reset_Click(object sender, RoutedEventArgs e)
        {
            _angle = 0;
            UpdateUI();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void Apply_Click(object sender, RoutedEventArgs e)
        {
            LeaderAngleSettings.AngleDegrees = _angle;
            LeaderAngleSettings.Save();
            DialogResult = true;
            Close();
        }

        private void UpdateUI()
        {
            _suppressTextChanged = true;
            txtCustom.Text = _angle > 0 ? _angle.ToString("0.##") : "";
            _suppressTextChanged = false;
            UpdatePresetHighlight();
            DrawPreview();
            UpdateLabel();
        }

        private void UpdatePresetHighlight()
        {
            foreach (var btn in _presetButtons)
            {
                var text = btn.Content.ToString().Replace("°", "");
                btn.Tag = double.TryParse(text, out double val) && Math.Abs(val - _angle) < 0.01 ? "Active" : null;
            }
        }

        private void UpdateLabel()
        {
            lblCurrent.Text = _angle > 0
                ? $"Current: {_angle:0.##}°"
                : "Current: Straight (default)";

            lblAngleInfo.Text = _angle > 0
                ? $"Elbow angle = {_angle:0.##}°"
                : "No elbow — straight leader";
        }

        private void DrawPreview()
        {
            previewCanvas.Children.Clear();

            double w = previewCanvas.ActualWidth > 0 ? previewCanvas.ActualWidth : 340;
            double h = previewCanvas.ActualHeight > 0 ? previewCanvas.ActualHeight : 200;

            int count = 5;
            double tagSpacing = h / (count + 1);

            // Elements spread horizontally at bottom
            double elemY = h * 0.85;
            double elemStartX = w * 0.10;
            double elemEndX = w * 0.65;
            double elemSpacing = (elemEndX - elemStartX) / (count - 1);

            // Tags stacked vertically on the right, evenly spaced
            double tagX = w * 0.88;
            double tagStartY = h * 0.08;

            var leaderBrush = new SolidColorBrush(Color.FromRgb(198, 125, 60));

            for (int i = 0; i < count; i++)
            {
                double eX = elemStartX + i * elemSpacing;
                double eY = elemY;
                double tX = tagX;
                double tY = tagStartY + i * tagSpacing;

                // Draw element (circle)
                DrawElement(eX, eY);

                // Draw tag head
                DrawTagHead(tX, tY);

                if (_angle <= 0)
                {
                    // Straight leader: element → tag
                    DrawLine(eX, eY, tX, tY, leaderBrush, 1.5);
                    DrawArrowTo(tX, tY, eX, eY, leaderBrush);
                }
                else
                {
                    // Leader: element → lên theo góc tới ngang bằng tag → ngang vào tag
                    double dy = eY - tY;
                    double elbowX;

                    if (Math.Abs(_angle - 90) < 0.01)
                    {
                        // 90°: thẳng đứng lên
                        elbowX = eX;
                    }
                    else
                    {
                        double angleRad = _angle * Math.PI / 180.0;
                        double offsetX = dy / Math.Tan(angleRad);
                        double maxOffsetX = Math.Abs(tX - eX);
                        elbowX = eX + Math.Sign(tX - eX) * Math.Min(offsetX, maxOffsetX);
                    }

                    // Đoạn 1: element → lên tới elbow (thẳng đứng hoặc xiên)
                    DrawLine(eX, eY, elbowX, tY, leaderBrush, 1.5);
                    // Đoạn 2: elbow → ngang vào tag
                    DrawLine(elbowX, tY, tX, tY, leaderBrush, 1.5);

                    // Elbow dot
                    var dot = new Ellipse { Width = 4, Height = 4, Fill = leaderBrush };
                    Canvas.SetLeft(dot, elbowX - 2);
                    Canvas.SetTop(dot, tY - 2);
                    previewCanvas.Children.Add(dot);

                    // Arrow at element
                    DrawArrowTo(elbowX, tY, eX, eY, leaderBrush);
                }
            }
        }

        private Point CalculateElbow(Point head, Point end, double angleDeg)
        {
            double angleRad = angleDeg * Math.PI / 180.0;
            double dy = end.Y - head.Y;
            double dx = end.X - head.X;

            if (Math.Abs(angleDeg - 90) < 0.01)
            {
                // 90° = go straight down then horizontal
                return new Point(head.X, end.Y);
            }

            // From head, go at angle until reaching end.Y level
            double elbowOffsetX = Math.Abs(dy) / Math.Tan(angleRad);
            double elbowX = head.X + Math.Sign(dx) * Math.Min(elbowOffsetX, Math.Abs(dx));

            return new Point(elbowX, end.Y);
        }

        private void DrawElement(double x, double y)
        {
            var circle = new Ellipse
            {
                Width = 20, Height = 20,
                Fill = new SolidColorBrush(Color.FromRgb(200, 200, 200)),
                Stroke = new SolidColorBrush(Color.FromRgb(150, 150, 150)),
                StrokeThickness = 1.5
            };
            Canvas.SetLeft(circle, x - 10);
            Canvas.SetTop(circle, y - 10);
            previewCanvas.Children.Add(circle);

            var label = new TextBlock
            {
                Text = "Element",
                FontSize = 10,
                Foreground = new SolidColorBrush(Color.FromRgb(120, 120, 120))
            };
            Canvas.SetLeft(label, x - 20);
            Canvas.SetTop(label, y + 12);
            previewCanvas.Children.Add(label);
        }

        private void DrawTagHead(double x, double y)
        {
            var rect = new Rectangle
            {
                Width = 44, Height = 18,
                RadiusX = 3, RadiusY = 3,
                Fill = new SolidColorBrush(Color.FromRgb(252, 243, 232)),
                Stroke = new SolidColorBrush(Color.FromRgb(198, 125, 60)),
                StrokeThickness = 1.2
            };
            Canvas.SetLeft(rect, x - 22);
            Canvas.SetTop(rect, y - 9);
            previewCanvas.Children.Add(rect);

            var label = new TextBlock
            {
                Text = "Tag",
                FontSize = 9,
                Foreground = new SolidColorBrush(Color.FromRgb(198, 125, 60)),
                FontWeight = FontWeights.SemiBold
            };
            Canvas.SetLeft(label, x - 10);
            Canvas.SetTop(label, y - 8);
            previewCanvas.Children.Add(label);
        }

        private void DrawLine(double x1, double y1, double x2, double y2, Brush stroke, double thickness)
        {
            previewCanvas.Children.Add(new Line
            {
                X1 = x1, Y1 = y1, X2 = x2, Y2 = y2,
                Stroke = stroke,
                StrokeThickness = thickness
            });
        }

        private void DrawArrowTo(double fromX, double fromY, double toX, double toY, Brush fill)
        {
            // Arrow pointing toward (toX, toY) from direction of last line segment
            double dx = toX - fromX;
            double dy = toY - fromY;
            double len = Math.Sqrt(dx * dx + dy * dy);
            if (len < 1) return;

            // Normalize
            double nx = dx / len;
            double ny = dy / len;

            // Arrow tip at element, size = 6px
            double tipX = toX;
            double tipY = toY;
            double baseX = tipX - nx * 8;
            double baseY = tipY - ny * 8;
            double perpX = -ny * 3.5;
            double perpY = nx * 3.5;

            var arrow = new Polygon
            {
                Points = new PointCollection
                {
                    new Point(tipX, tipY),
                    new Point(baseX + perpX, baseY + perpY),
                    new Point(baseX - perpX, baseY - perpY)
                },
                Fill = fill
            };
            previewCanvas.Children.Add(arrow);
        }
    }
}
