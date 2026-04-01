using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;
using THBIM.Models;

namespace THBIM.Helpers
{
    /// <summary>
    /// Tạo Border row cho một ParameterItem trong StackPanel Available/Selected.
    /// Tách riêng để tất cả các View dùng chung — không lặp code.
    /// </summary>
    public static class ParamRowHelper
    {
        // ── Màu sắc theo Kind ─────────────────────────────────────────────
        private static readonly SolidColorBrush BgInstance = new(Color.FromRgb(0xEA, 0xF3, 0xDE));
        private static readonly SolidColorBrush BgType = new(Color.FromRgb(0xFD, 0xF0, 0xCC));
        private static readonly SolidColorBrush BgReadOnly = new(Color.FromRgb(0xFD, 0xEA, 0xEA));
        private static readonly SolidColorBrush BgHighlight = new(Color.FromRgb(0xCC, 0xE5, 0xFF));
        private static readonly SolidColorBrush BgHover = new(Color.FromRgb(0xF0, 0xF6, 0xFF));

        private static readonly SolidColorBrush DotInstance = new(Color.FromRgb(0x7A, 0xB8, 0x7A));
        private static readonly SolidColorBrush DotType = new(Color.FromRgb(0xD4, 0xAA, 0x40));
        private static readonly SolidColorBrush DotReadOnly = new(Color.FromRgb(0xD0, 0x70, 0x70));

        private static readonly SolidColorBrush FgText = new(Color.FromRgb(0x1A, 0x1A, 0x1A));
        private static readonly SolidColorBrush BorderHl = new(Color.FromRgb(0x5B, 0x9B, 0xD5));

        // ── Tạo row Border ────────────────────────────────────────────────

        /// <summary>
        /// Tạo một Border row cho ParameterItem.
        /// Click vào row → toggle highlight (dùng để move ◀ ▶).
        /// Double-click → move ngay sang panel kia.
        /// </summary>
        /// <param name="param">ParameterItem cần hiển thị.</param>
        /// <param name="onDoubleClick">Action gọi khi double-click (thường là MoveRight/MoveLeft).</param>
        public static Border CreateRow(ParameterItem param,
                                       System.Action<ParameterItem> onDoubleClick = null)
        {
            var bg = GetBg(param.Kind);

            var border = new Border
            {
                Tag = param,
                Height = 22,
                Padding = new Thickness(6, 0, 6, 0),
                Background = bg,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand
            };

            // Inner layout: tên param bên trái, chấm tròn bên phải
            var sp = new StackPanel { Orientation = Orientation.Horizontal };

            var tbName = new TextBlock
            {
                Text = param.Name,
                FontSize = 12,
                Foreground = FgText,
                VerticalAlignment = VerticalAlignment.Center,
                Width = 200,
                TextTrimming = TextTrimming.CharacterEllipsis
            };

            var dot = new Ellipse
            {
                Width = 8,
                Height = 8,
                Fill = GetDot(param.Kind),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(4, 0, 0, 0)
            };

            sp.Children.Add(tbName);
            sp.Children.Add(dot);
            border.Child = sp;

            // ── Hover effect ──────────────────────────────────────────────
            border.MouseEnter += (s, e) =>
            {
                var b = (Border)s;
                var p = (ParameterItem)b.Tag;
                if (b.BorderThickness.Left < 1) // không hover nếu đang highlighted
                    b.Background = BgHover;
            };
            border.MouseLeave += (s, e) =>
            {
                var b = (Border)s;
                var p = (ParameterItem)b.Tag;
                b.Background = p.IsHighlighted ? BgHighlight : GetBg(p.Kind);
            };

            // ── Click → toggle highlight ──────────────────────────────────
            border.MouseLeftButtonDown += (s, e) =>
            {
                var b = (Border)s;
                var p = (ParameterItem)b.Tag;

                // Ctrl+click = multi-select; click thường = single select
                bool ctrl = Keyboard.IsKeyDown(Key.LeftCtrl) ||
                            Keyboard.IsKeyDown(Key.RightCtrl);

                if (!ctrl)
                {
                    // Bỏ highlight tất cả anh em
                    if (b.Parent is StackPanel parentSp)
                        foreach (Border sibling in parentSp.Children)
                            SetHighlight(sibling, false);
                }

                p.IsHighlighted = !p.IsHighlighted;
                SetHighlight(b, p.IsHighlighted);

                if (e.ClickCount == 2)
                    onDoubleClick?.Invoke(p);

                e.Handled = true;
            };

            return border;
        }

        // ── Toggle highlight trực tiếp trên Border ────────────────────────

        public static void SetHighlight(Border border, bool highlighted)
        {
            if (border.Tag is not ParameterItem p) return;
            p.IsHighlighted = highlighted;
            border.Background = highlighted ? BgHighlight : GetBg(p.Kind);
            border.BorderBrush = highlighted ? BorderHl : null;
            border.BorderThickness = highlighted
                ? new Thickness(2, 0, 0, 0)
                : new Thickness(0);
        }

        // ── Filter visibility ─────────────────────────────────────────────

        /// <summary>Lọc theo text search — ẩn/hiện các row trong StackPanel.</summary>
        public static void FilterByText(StackPanel panel, string query)
        {
            var q = (query ?? string.Empty).ToLower();
            foreach (Border row in panel.Children)
            {
                if (row.Tag is not ParameterItem p) continue;
                row.Visibility = string.IsNullOrEmpty(q) || p.Name.ToLower().Contains(q)
                    ? Visibility.Visible
                    : Visibility.Collapsed;
            }
        }

        /// <summary>Lọc theo Kind — "all" / "instance" / "type" / "readonly".</summary>
        public static void FilterByKind(StackPanel panel, string filter)
        {
            foreach (Border row in panel.Children)
            {
                if (row.Tag is not ParameterItem p) continue;
                row.Visibility = filter switch
                {
                    "instance" => p.Kind == ParamKind.Instance ? Visibility.Visible : Visibility.Collapsed,
                    "type" => p.Kind == ParamKind.Type ? Visibility.Visible : Visibility.Collapsed,
                    "readonly" => p.Kind == ParamKind.ReadOnly ? Visibility.Visible : Visibility.Collapsed,
                    _ => Visibility.Visible
                };
            }
        }

        // ── Private helpers ───────────────────────────────────────────────

        private static SolidColorBrush GetBg(ParamKind kind) => kind switch
        {
            ParamKind.Instance => BgInstance,
            ParamKind.Type => BgType,
            ParamKind.ReadOnly => BgReadOnly,
            _ => Brushes.White as SolidColorBrush
        };

        private static SolidColorBrush GetDot(ParamKind kind) => kind switch
        {
            ParamKind.Instance => DotInstance,
            ParamKind.Type => DotType,
            ParamKind.ReadOnly => DotReadOnly,
            _ => Brushes.Gray as SolidColorBrush
        };
    }
}