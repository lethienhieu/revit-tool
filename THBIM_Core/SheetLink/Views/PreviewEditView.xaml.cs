using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using THBIM.Models;

namespace THBIM
{
    public partial class PreviewEditView : UserControl
    {
        private Border _activeSheetTab;
        private IReadOnlyDictionary<string, ParamKind> _currentKindMap;

        // Color constants matching Excel export (ExcelService.cs)
        private static readonly SolidColorBrush BrushGreen  = new(Color.FromRgb(0xEA, 0xF3, 0xDE)); // Instance
        private static readonly SolidColorBrush BrushYellow = new(Color.FromRgb(0xFD, 0xF0, 0xCC)); // Type
        private static readonly SolidColorBrush BrushRed    = new(Color.FromRgb(0xFD, 0xEA, 0xEA)); // ReadOnly
        private static readonly SolidColorBrush BrushGray   = new(Color.FromRgb(0xF0, 0xF0, 0xF0)); // Element ID header
        private static readonly SolidColorBrush BrushTarget = new(Color.FromRgb(0xF7, 0xF7, 0xF7)); // Target column data
        private static readonly SolidColorBrush BrushBorder = new(Color.FromRgb(0xD0, 0xD0, 0xD0));

        // Light data cell tints (for data rows, lighter version of header)
        private static readonly SolidColorBrush BrushGreenLight  = new(Color.FromRgb(0xF5, 0xFA, 0xEF)); // Instance data
        private static readonly SolidColorBrush BrushYellowLight = new(Color.FromRgb(0xFE, 0xF8, 0xE6)); // Type data
        private static readonly SolidColorBrush BrushRedLight    = new(Color.FromRgb(0xFE, 0xF5, 0xF5)); // ReadOnly data

        public PreviewEditView()
        {
            InitializeComponent();
            Loaded += (_, _) =>
            {
                _activeSheetTab = TabRooms;
                SetSheetTabActive(TabRooms);
            };
        }

        // -- Progress bar --------------------------------------------------

        public void SetProgress(string message, bool success = true)
        {
            TbProgress.Text = message;
            ProgressFill.Background = success
                ? new SolidColorBrush(Color.FromRgb(0xC8, 0xEE, 0xC8))
                : new SolidColorBrush(Color.FromRgb(0xFD, 0xF0, 0xCC));
        }

        internal void LoadPreviewData(
            string sourceLabel,
            IReadOnlyList<string> targets,
            IReadOnlyList<string> parameters,
            IReadOnlyList<PreviewRowData> rows = null,
            IReadOnlyDictionary<string, ParamKind> paramKindMap = null)
        {
            _currentKindMap = paramKindMap ?? new Dictionary<string, ParamKind>(System.StringComparer.OrdinalIgnoreCase);

            var targetList = targets?
                .Where(t => !string.IsNullOrWhiteSpace(t))
                .Distinct(System.StringComparer.OrdinalIgnoreCase)
                .ToList() ?? new List<string>();
            var parameterList = parameters?
                .Where(p => !string.IsNullOrWhiteSpace(p))
                .Distinct(System.StringComparer.OrdinalIgnoreCase)
                .ToList() ?? new List<string>();
            var rowList = rows?
                .Where(r => r != null)
                .Take(PreviewValueHelpers.MaxRows)
                .ToList() ?? new List<PreviewRowData>();

            if (!rowList.Any())
            {
                rowList = targetList
                    .Take(PreviewValueHelpers.MaxRows)
                    .Select(t => new PreviewRowData(t, new Dictionary<string, string>()))
                    .ToList();
            }

            if (_activeSheetTab == null)
            {
                _activeSheetTab = TabRooms;
                SetSheetTabActive(TabRooms);
            }

            if (TabRooms?.Child is TextBlock tabText)
                tabText.Text = string.IsNullOrWhiteSpace(sourceLabel) ? "Preview" : sourceLabel;

            if (!targetList.Any() || !parameterList.Any())
            {
                SetProgress("No preview data yet. Select objects and parameters to enable Preview.", false);
                RenderPreviewMatrix(new List<PreviewRowData>(), new List<string>());
                return;
            }

            var parameterPreview = string.Join(", ", parameterList.Take(3));
            if (parameterList.Count > 3) parameterPreview += $" +{parameterList.Count - 3}";

            SetProgress(
                $"{sourceLabel}: {targetList.Count} objects | {parameterList.Count} parameters | {parameterPreview}",
                true);

            RenderPreviewMatrix(rowList, parameterList);
        }

        // -- Sheet tabs ----------------------------------------------------

        private void SheetTab_Click(object sender, MouseButtonEventArgs e)
        {
            if (sender is Border tab) SetSheetTabActive(tab);
        }

        private void SetSheetTabActive(Border tab)
        {
            if (_activeSheetTab != null)
            {
                _activeSheetTab.Background = new SolidColorBrush(Color.FromRgb(0xD0, 0xD0, 0xD0));
                if (_activeSheetTab.Child is TextBlock prev)
                {
                    prev.FontWeight = FontWeights.Normal;
                    prev.Foreground = new SolidColorBrush(Color.FromRgb(0x55, 0x55, 0x55));
                }
            }

            _activeSheetTab = tab;
            tab.Background = Brushes.White;
            if (tab.Child is TextBlock tb)
            {
                tb.FontWeight = FontWeights.Medium;
                tb.Foreground = new SolidColorBrush(Color.FromRgb(0x1A, 0x1A, 0x1A));
            }

            var showInstructions = string.Equals(tab.Tag?.ToString(), "Instructions", System.StringComparison.OrdinalIgnoreCase);
            RoomsSheetScroll.Visibility = showInstructions ? Visibility.Collapsed : Visibility.Visible;
            InstructionsPanel.Visibility = showInstructions ? Visibility.Visible : Visibility.Collapsed;
        }

        private void BtnSheetLeft_Click(object sender, RoutedEventArgs e) { }
        private void BtnSheetRight_Click(object sender, RoutedEventArgs e) { }

        private void BtnAddSheet_Click(object sender, MouseButtonEventArgs e)
            => MessageBox.Show("Add sheet — coming soon.",
               "SheetLink", MessageBoxButton.OK, MessageBoxImage.Information);

        // -- Action bar ----------------------------------------------------

        private void BtnReset_Click(object sender, RoutedEventArgs e)
            => SetProgress("Completed   0%");

        private void BtnUpdateModel_Click(object sender, RoutedEventArgs e)
        {
            var res = MessageBox.Show(
                "Update data back to Revit?\nThis action cannot be undone.",
                "SheetLink", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (res != MessageBoxResult.Yes) return;

            SetProgress("Updating model...", false);
            MessageBox.Show("Update Model — Phase 3 (Revit API).",
                "SheetLink", MessageBoxButton.OK, MessageBoxImage.Information);
            SetProgress("[1/1] Completed   100%", true);
        }

        // -- Preview matrix ------------------------------------------------

        private ParamKind GetParamKind(string paramName)
        {
            if (_currentKindMap != null && !string.IsNullOrWhiteSpace(paramName)
                && _currentKindMap.TryGetValue(paramName, out var kind))
                return kind;
            return ParamKind.Instance; // default
        }

        private static SolidColorBrush GetHeaderBrush(ParamKind kind) => kind switch
        {
            ParamKind.Instance => BrushGreen,
            ParamKind.Type     => BrushYellow,
            ParamKind.ReadOnly => BrushRed,
            _                  => BrushGreen
        };

        private static SolidColorBrush GetDataBrush(ParamKind kind) => kind switch
        {
            ParamKind.Instance => BrushGreenLight,
            ParamKind.Type     => BrushYellowLight,
            ParamKind.ReadOnly => BrushRedLight,
            _                  => BrushGreenLight
        };

        private void RenderPreviewMatrix(IReadOnlyList<PreviewRowData> rows, IReadOnlyList<string> parameters)
        {
            var rowList = rows?.Where(r => r != null).ToList() ?? new List<PreviewRowData>();
            var paramList = parameters?
                .Where(p => !string.IsNullOrWhiteSpace(p))
                .Distinct(System.StringComparer.OrdinalIgnoreCase)
                .ToList() ?? new List<string>();

            var grid = new Grid
            {
                Background = Brushes.White
            };

            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(240) });
            foreach (var _ in paramList)
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(170) });

            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(34) });
            foreach (var _ in rowList)
                grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(24) });

            // Header row — Element ID column
            AddCell(grid, 0, 0, "Element ID", isHeader: true, background: BrushGray);

            // Header row — parameter columns with kind-based coloring
            for (int c = 0; c < paramList.Count; c++)
            {
                var kind = GetParamKind(paramList[c]);
                AddCell(grid, 0, c + 1, paramList[c], isHeader: true, background: GetHeaderBrush(kind));
            }

            // Data rows
            for (int r = 0; r < rowList.Count; r++)
            {
                var row = rowList[r];
                AddCell(grid, r + 1, 0, row.Target, isHeader: false, background: BrushTarget);

                for (int c = 0; c < paramList.Count; c++)
                {
                    var val = row.GetValue(paramList[c]);
                    var kind = GetParamKind(paramList[c]);
                    AddCell(grid, r + 1, c + 1, val, isHeader: false, background: GetDataBrush(kind));
                }
            }

            RoomsSheetScroll.Content = grid;
        }

        private static void AddCell(Grid grid, int row, int column, string text, bool isHeader, SolidColorBrush background)
        {
            var border = new Border
            {
                BorderBrush = BrushBorder,
                BorderThickness = new Thickness(0, 0, 1, 1),
                Background = background,
                Padding = new Thickness(6, 2, 6, 2)
            };

            border.Child = new TextBlock
            {
                Text = text ?? string.Empty,
                FontSize = isHeader ? 12 : 11.5,
                FontWeight = isHeader ? FontWeights.Medium : FontWeights.Normal,
                Foreground = new SolidColorBrush(Color.FromRgb(0x2A, 0x2A, 0x2A)),
                VerticalAlignment = VerticalAlignment.Center,
                TextTrimming = TextTrimming.CharacterEllipsis
            };

            Grid.SetRow(border, row);
            Grid.SetColumn(border, column);
            grid.Children.Add(border);
        }
    }
}
