using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using THBIM.Helpers;
using THBIM.Models;
using THBIM.Services;

namespace THBIM
{
    public partial class SpatialView : UserControl
    {
        private readonly MockDataService _mock = new();
        private bool _isRooms = true;
        private List<SpatialItem> _currentItems = new();

        public SpatialView() { InitializeComponent(); Loaded += (_, _) => LoadData(); }

        // ── Load ──────────────────────────────────────────────────────────

        private void LoadData()
        {
            try
            {
                _currentItems = _isRooms
                    ? (ServiceLocator.IsRevitMode ? ServiceLocator.RevitData.GetRooms() : _mock.GetRooms())
                    : (ServiceLocator.IsRevitMode ? ServiceLocator.RevitData.GetSpaces() : _mock.GetSpaces());
            }
            catch
            {
                _currentItems = _isRooms ? _mock.GetRooms() : _mock.GetSpaces();
            }

            SpSpatial.Children.Clear();
            foreach (var item in _currentItems)
                SpSpatial.Children.Add(BuildSpatialRow(item));

            List<ParameterItem> parms;
            try
            {
                parms = ServiceLocator.IsRevitMode
                    ? ServiceLocator.RevitData.GetSpatialParameters(_isRooms)
                    : _mock.GetSpatialParameters();
            }
            catch
            {
                parms = new List<ParameterItem>();
            }
            if (!parms.Any())
                parms = _mock.GetSpatialParameters();

            SpAvail.Children.Clear();
            foreach (var p in parms)
                SpAvail.Children.Add(CreateAvailableRow(p));
            
            TbSelectHeader.Text = _isRooms ? "Select Rooms" : "Select Spaces";
            ApplyAvailableFilters();
            UpdateStatus();
        }

        private Border BuildSpatialRow(SpatialItem item)
        {
            var row = new Border
            {
                Tag = item,
                Height = 22,
                Background = Brushes.Transparent,
                Cursor = Cursors.Hand
            };

            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(32) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(80) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var cb = new CheckBox
            {
                IsChecked = item.IsChecked,
                Tag = item,
                Margin = new Thickness(0),
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalContentAlignment = VerticalAlignment.Center
            };
            cb.Checked += SpatialCheckBox_Click;
            cb.Unchecked += SpatialCheckBox_Click;
            Grid.SetColumn(cb, 0);

            var tbNum = new TextBlock
            {
                Text = item.Number,
                FontSize = 12,
                Foreground = new SolidColorBrush(Color.FromRgb(0x55, 0x55, 0x55)),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(4, 0, 0, 0)
            };
            Grid.SetColumn(tbNum, 1);

            var div = new Border { Background = new SolidColorBrush(Color.FromRgb(0xE0, 0xE0, 0xE0)) };
            Grid.SetColumn(div, 2);

            var tbName = new TextBlock
            {
                Text = item.Name,
                FontSize = 12,
                Foreground = new SolidColorBrush(Color.FromRgb(0x1A, 0x1A, 0x1A)),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(8, 0, 0, 0),
                TextTrimming = TextTrimming.CharacterEllipsis
            };
            Grid.SetColumn(tbName, 3);

            grid.Children.Add(cb);
            grid.Children.Add(tbNum);
            grid.Children.Add(div);
            grid.Children.Add(tbName);
            row.Child = grid;

            row.MouseEnter += (s, _) => ((Border)s).Background = new SolidColorBrush(Color.FromRgb(0xF0, 0xF6, 0xFF));
            row.MouseLeave += (s, _) => ((Border)s).Background = Brushes.Transparent;
            row.MouseLeftButtonDown += SpatialRow_Click;

            return row;
        }

        // ── Events ────────────────────────────────────────────────────────

        private void SpatialRow_Click(object sender, MouseButtonEventArgs e)
        {
            if (sender is not Border row) return;
            var cb = (row.Child as Grid)?.Children.OfType<CheckBox>().FirstOrDefault();
            if (cb != null) cb.IsChecked = cb.IsChecked != true;
        }

        private void SpatialCheckBox_Click(object sender, RoutedEventArgs e)
        {
            if (sender is CheckBox cb && cb.Tag is SpatialItem item)
                item.IsChecked = cb.IsChecked == true;

            int chk = _currentItems.Count(i => i.IsChecked);
            ChkAllSpatial.IsChecked = chk == _currentItems.Count ? true : chk == 0 ? false : (bool?)null;
            UpdateStatus();
        }

        private void ChkAllSpatial_Click(object sender, RoutedEventArgs e)
        {
            bool check = ChkAllSpatial.IsChecked == true;
            foreach (var item in _currentItems) item.IsChecked = check;
            foreach (Border row in SpSpatial.Children)
            {
                var cb = (row.Child as Grid)?.Children.OfType<CheckBox>().FirstOrDefault();
                if (cb != null) cb.IsChecked = check;
            }
            UpdateStatus();
        }

        // ── Search ────────────────────────────────────────────────────────

        private void TbSrchSpatial_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (SpSpatial == null) return;

            var q = TbSrchSpatial.Text.ToLower();
            foreach (Border row in SpSpatial.Children)
            {
                if (row.Tag is not SpatialItem item) continue;
                row.Visibility = string.IsNullOrEmpty(q) ||
                                 item.Number.ToLower().Contains(q) ||
                                 item.Name.ToLower().Contains(q)
                    ? Visibility.Visible : Visibility.Collapsed;
            }
        }

        private void TbSrchAvail_TextChanged(object sender, TextChangedEventArgs e)
        {
            ApplyAvailableFilters();
        }
        private void TbSrchSel_TextChanged(object sender, TextChangedEventArgs e)
        {
            ApplySelectedFilters();
        }

        private void ChkHide_Click(object sender, RoutedEventArgs e)
        {
            bool hide = ChkHide.IsChecked == true;
            foreach (Border row in SpSpatial.Children)
            {
                if (row.Tag is not SpatialItem item) continue;
                row.Visibility = (!hide || item.IsChecked) ? Visibility.Visible : Visibility.Collapsed;
            }
        }

        // ── Type switch / Filter ──────────────────────────────────────────

        private void CbSpatialType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!IsLoaded || SpSpatial == null || SpSel == null) return;
            if (CbSpatialType.SelectedItem is not ComboBoxItem item) return;
            _isRooms = item.Content?.ToString() == "Rooms";
            SpSel.Children.Clear();
            LoadData();
        }

        private void BtnRefresh_Click(object sender, RoutedEventArgs e) => LoadData();

        private void BadgeInst_Click(object sender, MouseButtonEventArgs e)
        { ChkAvailInstance.IsChecked = true; ChkAvailReadOnly.IsChecked = false; ParamFilter_CheckedChanged(sender, e); }
        private void BadgeRo_Click(object sender, MouseButtonEventArgs e)
        { ChkAvailInstance.IsChecked = false; ChkAvailReadOnly.IsChecked = true; ParamFilter_CheckedChanged(sender, e); }

        private void ParamFilter_CheckedChanged(object sender, RoutedEventArgs e)
        {
            EnsureAtLeastOneChecked(ChkAvailInstance, ChkAvailReadOnly);
            UpdateParamFilterText();
            ApplyAvailableFilters();
        }

        private void UpdateParamFilterText()
        {
            var parts = new List<string>();
            if (ChkAvailInstance?.IsChecked == true) parts.Add("Instance");
            if (ChkAvailReadOnly?.IsChecked == true) parts.Add("Read-only");
            if (CbParamFilter != null) CbParamFilter.Text = string.Join(", ", parts);
        }

        private void SelParamFilter_CheckedChanged(object sender, RoutedEventArgs e)
        {
            EnsureAtLeastOneChecked(ChkSelInstance, ChkSelReadOnly);
            UpdateSelParamFilterText();
            ApplySelectedFilters();
        }

        private void UpdateSelParamFilterText()
        {
            var parts = new List<string>();
            if (ChkSelInstance?.IsChecked == true) parts.Add("Instance");
            if (ChkSelReadOnly?.IsChecked == true) parts.Add("Read-only");
            if (CbSelParamFilter != null) CbSelParamFilter.Text = string.Join(", ", parts);
        }

        private static void EnsureAtLeastOneChecked(CheckBox a, CheckBox b)
        {
            if (a.IsChecked == true || b.IsChecked == true) return;
            a.IsChecked = true;
            b.IsChecked = true;
        }

        private static bool PassKind(ParamKind kind, CheckBox inst, CheckBox readOnly)
            => kind switch
            {
                ParamKind.Instance => inst.IsChecked == true,
                ParamKind.Type => false, // No type params in spatial view
                ParamKind.ReadOnly => readOnly.IsChecked == true,
                _ => false
            };

        private void ApplyAvailableFilters()
        {
            if (SpAvail == null) return;
            var q = TbSrchAvail.Text?.Trim() ?? string.Empty;
            foreach (var row in SpAvail.Children.OfType<Border>())
            {
                if (row.Tag is not ParameterItem p) continue;
                var pass = PassKind(p.Kind, ChkAvailInstance, ChkAvailReadOnly) &&
                           (string.IsNullOrWhiteSpace(q) || p.Name.IndexOf(q, System.StringComparison.OrdinalIgnoreCase) >= 0);
                row.Visibility = pass ? Visibility.Visible : Visibility.Collapsed;
            }
            UpdateStatus();
        }

        private void ApplySelectedFilters()
        {
            if (SpSel == null) return;
            var q = TbSrchSel.Text?.Trim() ?? string.Empty;
            foreach (var row in SpSel.Children.OfType<Border>())
            {
                if (row.Tag is not ParameterItem p) continue;
                var passKind = PassKind(p.Kind, ChkSelInstance, ChkSelReadOnly);
                var passText = string.IsNullOrWhiteSpace(q) || p.Name.IndexOf(q, System.StringComparison.OrdinalIgnoreCase) >= 0;
                row.Visibility = passKind && passText ? Visibility.Visible : Visibility.Collapsed;
            }
            UpdateStatus();
        }

        // ── Move ──────────────────────────────────────────────────────────

        private void BtnRight_Click(object sender, RoutedEventArgs e)
        {
            var toMove = SpAvail.Children.OfType<Border>()
                .Where(b => b.Visibility == Visibility.Visible && b.Tag is ParameterItem p && p.IsHighlighted)
                .Select(b => b.Tag as ParameterItem).Where(p => p != null).ToList();
            if (!toMove.Any())
                toMove = SpAvail.Children.OfType<Border>()
                    .Where(b => b.Visibility == Visibility.Visible)
                    .Select(b => b.Tag as ParameterItem).Where(p => p != null).ToList();
            foreach (var p in toMove) MoveToSelected(p);
        }

        private void BtnLeft_Click(object sender, RoutedEventArgs e)
        {
            var toMove = SpSel.Children.OfType<Border>()
                .Where(b => b.Visibility == Visibility.Visible && b.Tag is ParameterItem p && p.IsHighlighted)
                .Select(b => b.Tag as ParameterItem).Where(p => p != null).ToList();
            if (!toMove.Any())
                toMove = SpSel.Children.OfType<Border>()
                    .Where(b => b.Visibility == Visibility.Visible)
                    .Select(b => b.Tag as ParameterItem).Where(p => p != null).ToList();
            foreach (var p in toMove) MoveToAvailable(p);
        }

        private void MoveToSelected(ParameterItem p)
        {
            if (p == null || SpSel.Children.OfType<Border>().Any(b => SameParam(b.Tag as ParameterItem, p))) return;
            var row = SpAvail.Children.OfType<Border>().FirstOrDefault(b => SameParam(b.Tag as ParameterItem, p));
            if (row != null) SpAvail.Children.Remove(row);
            p.IsHighlighted = false;
            SpSel.Children.Add(CreateSelectedRow(p));
            ApplyAvailableFilters();
            ApplySelectedFilters();
        }

        private void MoveToAvailable(ParameterItem p)
        {
            if (p == null || SpAvail.Children.OfType<Border>().Any(b => SameParam(b.Tag as ParameterItem, p))) return;
            var row = SpSel.Children.OfType<Border>().FirstOrDefault(b => SameParam(b.Tag as ParameterItem, p));
            if (row != null) SpSel.Children.Remove(row);
            p.IsHighlighted = false;
            SpAvail.Children.Add(CreateAvailableRow(p));
            ApplyAvailableFilters();
            ApplySelectedFilters();
        }

        private Border CreateAvailableRow(ParameterItem p) => ParamRowHelper.CreateRow(p, MoveToSelected);
        private Border CreateSelectedRow(ParameterItem p) => ParamRowHelper.CreateRow(p, MoveToAvailable);
        private static bool SameParam(ParameterItem a, ParameterItem b) => a != null && b != null && string.Equals(a.Name, b.Name, System.StringComparison.OrdinalIgnoreCase);

        // ── Sort ──────────────────────────────────────────────────────────

        private Border GetHlSel()
            => SpSel.Children.OfType<Border>()
                    .FirstOrDefault(b => b.Tag is ParameterItem p && p.IsHighlighted);

        private void SrtTop_Click(object s, RoutedEventArgs e) { var r = GetHlSel(); if (r == null) return; SpSel.Children.Remove(r); SpSel.Children.Insert(0, r); }
        private void SrtUp_Click(object s, RoutedEventArgs e) { var r = GetHlSel(); if (r == null) return; int i = SpSel.Children.IndexOf(r); if (i > 0) { SpSel.Children.Remove(r); SpSel.Children.Insert(i - 1, r); } }
        private void SrtDown_Click(object s, RoutedEventArgs e) { var r = GetHlSel(); if (r == null) return; int i = SpSel.Children.IndexOf(r); if (i < SpSel.Children.Count - 1) { SpSel.Children.Remove(r); SpSel.Children.Insert(i + 1, r); } }
        private void SrtBot_Click(object s, RoutedEventArgs e) { var r = GetHlSel(); if (r == null) return; SpSel.Children.Remove(r); SpSel.Children.Add(r); }
        private void SrtReset_Click(object s, RoutedEventArgs e)
        {
            var items = SpSel.Children.OfType<Border>().OrderBy(b => (b.Tag as ParameterItem)?.Name, System.StringComparer.OrdinalIgnoreCase)
                             .OrderBy(b => (b.Tag as ParameterItem)?.Name).ToList();
            SpSel.Children.Clear();
            foreach (var i in items) SpSel.Children.Add(i);
        }

        // ── Action bar ────────────────────────────────────────────────────

        private void BtnReset_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in _currentItems) item.IsChecked = false;
            foreach (Border row in SpSpatial.Children)
            {
                var cb = (row.Child as Grid)?.Children.OfType<CheckBox>().FirstOrDefault();
                if (cb != null) cb.IsChecked = false;
            }
            ChkAllSpatial.IsChecked = false;
            SpSel.Children.Clear();

            ChkAvailInstance.IsChecked = true;
            ChkAvailReadOnly.IsChecked = true;
            UpdateParamFilterText();

            ChkSelInstance.IsChecked = true;
            ChkSelReadOnly.IsChecked = true;
            UpdateSelParamFilterText();

            TbSrchSpatial.Clear(); TbSrchAvail.Clear(); TbSrchSel.Clear();
            LoadData();
        }

        private void BtnPreview_Click(object sender, RoutedEventArgs e)
            => NavigateToPreview();

        private void BtnImport_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            { Filter = "Excel Files|*.xlsx;*.xls|All Files|*.*" };
            if (dlg.ShowDialog() != true) return;
            if (!ServiceLocator.IsRevitMode) { ShowImportResult(new ImportResult(), "Import Spatial"); return; }

            var handler = Services.RevitEventHandler.Instance;
            if (handler == null) { MessageBox.Show("Revit event handler not initialized.", "THBIM", MessageBoxButton.OK, MessageBoxImage.Error); return; }

            var host = Window.GetWindow(this) as SheetLinkWindow;
            host?.ShowBusy("Importing Spatial data...");

            ImportResult result = null;
            handler.Enqueue(
                _ =>
                {
                    result = ServiceLocator.Excel.ImportSpatialFromExcel(
                        dlg.FileName, RevitDocumentCache.Current, _isRooms,
                        (__, msg) => Dispatcher.Invoke(() => host?.UpdateBusyMessage(msg)));
                },
                () => Dispatcher.Invoke(() =>
                {
                    host?.HideBusy();
                    if (result != null) { ShowImportResult(result, "Import Spatial"); LoadData(); }
                }));
        }

        private void CbImportAction_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!IsLoaded || sender is not ComboBox cb || cb.SelectedIndex <= 0)
                return;

            BtnImport_Click(sender, new RoutedEventArgs());
            cb.SelectedIndex = 0;
        }

        private async void BtnExportTemplate_Click(object sender, RoutedEventArgs e)
        {
            var ids = _currentItems.Where(i => i.IsChecked).Select(i => i.ElementId).ToList();
            var parms = SpSel.Children.OfType<Border>()
                             .Select(b => (b.Tag as ParameterItem)?.Name)
                             .Where(n => n != null).ToList();

            if (!ids.Any() || !parms.Any())
            {
                MessageBox.Show("Select a Room/Space and a Parameter.", "SheetLink",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var opts = ShowExportDialog();
            if (opts == null) return;
            if (opts.Google)
            {
                MessageBox.Show("Google export will be implemented later.", "THBIM", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = $"SheetLink_{(_isRooms ? "Rooms" : "Spaces")}_{System.DateTime.Now:yyyyMMdd}.xlsx"
            };
            if (dlg.ShowDialog() != true) return;

            await RunWithBusy("Exporting...", async () =>
            {
                if (ServiceLocator.IsRevitMode)
                    ServiceLocator.Excel.ExportSpatial(
                        dlg.FileName, _isRooms, ids, parms);
                await Task.CompletedTask;
            }, "Export complete!");

            if (opts.OpenAfter)
                OpenFile(dlg.FileName);
        }

        // ── Status ────────────────────────────────────────────────────────

        private void UpdateStatus()
        {
            int sel = _currentItems.Count(i => i.IsChecked);
            int avail = SpAvail.Children.OfType<Border>().Count(b => b.Visibility == Visibility.Visible);
            int selP = SpSel.Children.Count;
            TbStatus.Text = $"Total number of selected {(_isRooms ? "rooms" : "spaces")} {sel} | parameters found {avail} | parameters selected {selP}";
            BtnPreview.IsEnabled = GetPreviewParameters().Any();
        }
    }
}
