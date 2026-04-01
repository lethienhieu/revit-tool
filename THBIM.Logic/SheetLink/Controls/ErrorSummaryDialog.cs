using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using THBIM.Services;

namespace THBIM.Controls
{
    public class ErrorSummaryDialog : Window
    {
        private readonly ImportResult _result;
        private readonly string       _op;

        public ErrorSummaryDialog(ImportResult result, string operationName = "Import")
        {
            _result = result;
            _op     = operationName;
            Title                 = $"SheetLink — {operationName} Result";
            Width                 = 560; Height = 420;
            MinWidth              = 400; MinHeight = 300;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ResizeMode            = ResizeMode.CanResize;
            WindowStyle           = WindowStyle.ToolWindow;
            Background            = new SolidColorBrush(Color.FromRgb(0xF4, 0xF4, 0xF4));
            FontFamily            = new FontFamily("Segoe UI");
            ShowInTaskbar         = false;
            Build();
        }

        private void Build()
        {
            var root = new Grid { Margin = new Thickness(16) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(8) });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(8) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            bool hasErr = _result.Errors.Any();
            var bg  = hasErr ? Color.FromRgb(0xFD,0xEA,0xEA) : Color.FromRgb(0xEA,0xF3,0xDE);
            var fg  = hasErr ? Color.FromRgb(0x92,0x22,0x22) : Color.FromRgb(0x2A,0x6E,0x2A);
            var bdr = hasErr ? Color.FromRgb(0xE8,0xA0,0xA0) : Color.FromRgb(0xB5,0xD9,0x8A);

            var summaryBorder = new Border
            {
                Background = new SolidColorBrush(bg),
                BorderBrush = new SolidColorBrush(bdr),
                BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(6),
                Padding = new Thickness(12, 8, 12, 8)
            };
            var sp = new StackPanel { Orientation = Orientation.Horizontal };
            sp.Children.Add(new TextBlock
            {
                Text = hasErr ? "⚠ " : "✓ ", FontSize = 16,
                Foreground = new SolidColorBrush(fg), VerticalAlignment = VerticalAlignment.Center
            });
            var txt = new StackPanel { Orientation = Orientation.Vertical };
            txt.Children.Add(new TextBlock
            {
                Text = hasErr ? $"{_op} completed with {_result.Errors.Count} error(s)" : $"{_op} succeeded",
                FontSize = 13, FontWeight = FontWeights.Medium, Foreground = new SolidColorBrush(fg)
            });
            txt.Children.Add(new TextBlock
            {
                Text = $"Parameters updated: {_result.Updated}",
                FontSize = 11.5, Foreground = new SolidColorBrush(fg), Opacity = 0.8
            });
            sp.Children.Add(txt);
            summaryBorder.Child = sp;
            Grid.SetRow(summaryBorder, 0);
            root.Children.Add(summaryBorder);

            if (hasErr)
            {
                var errPanel = new Grid();
                errPanel.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                errPanel.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
                Grid.SetRow(errPanel, 2);

                var hdr = new TextBlock
                {
                    Text = $"Error details ({_result.Errors.Count}):", FontSize = 12,
                    FontWeight = FontWeights.Medium,
                    Foreground = new SolidColorBrush(Color.FromRgb(0x55,0x55,0x55)),
                    Margin = new Thickness(0,0,0,4)
                };
                Grid.SetRow(hdr, 0);
                errPanel.Children.Add(hdr);

                var listBorder = new Border
                {
                    BorderBrush = new SolidColorBrush(Color.FromRgb(0xD0,0xD0,0xD0)),
                    BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(4),
                    Background = Brushes.White
                };
                var scroll = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
                var list   = new StackPanel { Margin = new Thickness(8, 6, 8, 6) };
                foreach (var err in _result.Errors)
                    list.Children.Add(new TextBlock
                    {
                        Text = $"• {err}", FontSize = 11.5, TextWrapping = TextWrapping.Wrap,
                        Foreground = new SolidColorBrush(Color.FromRgb(0x70,0x20,0x20)),
                        Margin = new Thickness(0, 1, 0, 1)
                    });
                scroll.Content = list;
                listBorder.Child = scroll;
                Grid.SetRow(listBorder, 1);
                errPanel.Children.Add(listBorder);
                root.Children.Add(errPanel);
            }
            else
            {
                var ok = new TextBlock
                {
                    Text = "✓ No errors.", FontSize = 13,
                    Foreground = new SolidColorBrush(Color.FromRgb(0x2A,0x6E,0x2A)),
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment   = VerticalAlignment.Center
                };
                Grid.SetRow(ok, 2);
                root.Children.Add(ok);
            }

            var btnRow = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };
            Grid.SetRow(btnRow, 4);

            if (hasErr)
            {
                var btnCopy = new Button
                {
                    Content = "Copy Errors", Width = 90, Height = 30,
                    Margin = new Thickness(0,0,8,0),
                    Background = new SolidColorBrush(Color.FromRgb(0xE0,0xE0,0xE0)),
                    BorderBrush = new SolidColorBrush(Color.FromRgb(0xD0,0xD0,0xD0)),
                    BorderThickness = new Thickness(1), FontSize = 12,
                    Cursor = System.Windows.Input.Cursors.Hand
                };
                btnCopy.Click += (_, _) =>
                {
                    Clipboard.SetText(string.Join("\n", _result.Errors));
                    btnCopy.Content = "✓ Copied!";
                };
                btnRow.Children.Add(btnCopy);
            }

            var btnClose = new Button
            {
                Content = "Close", Width = 80, Height = 30,
                IsDefault = true, IsCancel = true,
                Background = new SolidColorBrush(Color.FromRgb(0x2A,0x6E,0x2A)),
                Foreground = Brushes.White, BorderThickness = new Thickness(0),
                FontSize = 12, FontWeight = FontWeights.Medium,
                Cursor = System.Windows.Input.Cursors.Hand
            };
            btnClose.Click += (_, _) => Close();
            btnRow.Children.Add(btnClose);
            root.Children.Add(btnRow);
            Content = root;
        }

        public static void Show(ImportResult result, string operationName = "Import",
                                  Window owner = null)
        {
            Application.Current?.Dispatcher.Invoke(() =>
            {
                var dlg = new ErrorSummaryDialog(result, operationName);
                if (owner != null) dlg.Owner = owner;
                dlg.ShowDialog();
            });
        }
    }
}
