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
    public partial class ElementsView : UserControl
    {
        private readonly MockDataService _mock = new();
        private string _selCatName;
        private Border _selCatRow;

        private static readonly SolidColorBrush BgSel = new(Color.FromRgb(0xDD, 0xEE, 0xFF));
        private static readonly SolidColorBrush BgNorm = System.Windows.Media.Brushes.Transparent;
        private static readonly SolidColorBrush BgHover = new(Color.FromRgb(0xF0, 0xF6, 0xFF));

        private List<CategoryItem> _allCats = new();

        public ElementsView() { InitializeComponent(); Loaded += (_, _) => LoadCats(); }

        // ── Load categories ───────────────────────────────────────────────

        private void LoadCats()
        {
            _allCats = ServiceLocator.IsRevitMode
                ? ServiceLocator.RevitData.GetModelCategories()
                    .Concat(ServiceLocator.RevitData.GetAnnotationCategories())
                    .OrderBy(c => c.Name).ToList()
                : _mock.GetModelCategories()
                    .Concat(_mock.GetAnnotationCategories())
                    .OrderBy(c => c.Name).ToList();

            NormalizeDisciplineSelection();
            ApplyDisciplineFilter();
        }

        // ── Category single-select ────────────────────────────────────────

        private void Cat_Click(object sender, MouseButtonEventArgs e)
        {
            if (sender is not Border row || row.Tag is not CategoryItem cat) return;

            if (_selCatRow != null) _selCatRow.Background = BgNorm;
            _selCatRow = row;
            row.Background = BgSel;
            _selCatName = cat.Name;

            // Load element types
            SpElm.Children.Clear();
            var elms = ServiceLocator.IsRevitMode
                ? ServiceLocator.RevitData.GetElementTypes(cat.Name)
                : _mock.GetElements(cat.Name);

            foreach (var el in elms)
            {
                var r = new Border
                {
                    Tag = el,
                    Height = 22,
                    Padding = new Thickness(8, 0, 8, 0),
                    Background = BgNorm,
                    Cursor = Cursors.Hand
                };
                var sp = new StackPanel { Orientation = Orientation.Horizontal };
                var cb = new CheckBox
                {
                    IsChecked = el.IsChecked,
                    Tag = el,
                    Margin = new Thickness(0, 0, 6, 0),
                    VerticalContentAlignment = VerticalAlignment.Center
                };
                cb.Checked += ElmCb;
                cb.Unchecked += ElmCb;
                sp.Children.Add(cb);
                sp.Children.Add(new TextBlock
                {
                    Text = el.Name,
                    FontSize = 12,
                    Foreground = new SolidColorBrush(Color.FromRgb(0x1A, 0x1A, 0x1A)),
                    VerticalAlignment = VerticalAlignment.Center
                });
                r.Child = sp;
                r.MouseEnter += (s, _) => ((Border)s).Background = BgHover;
                r.MouseLeave += (s, _) => ((Border)s).Background = BgNorm;
                SpElm.Children.Add(r);
            }

            LoadParamsForCategory(cat.Name);

            UpdateStatus();
        }

        // ── Element checkbox ──────────────────────────────────────────────

        private void ElmCb(object sender, RoutedEventArgs e)
        {
            if (sender is CheckBox cb && cb.Tag is CategoryItem el)
                el.IsChecked = cb.IsChecked == true;
            UpdateStatus();
        }

        private void ChkAllElm_Click(object sender, RoutedEventArgs e)
        {
            bool check = ChkAllElm.IsChecked == true;
            foreach (Border row in SpElm.Children)
            {
                var cb = (row.Child as StackPanel)?.Children.OfType<CheckBox>().FirstOrDefault();
                if (cb != null) { cb.IsChecked = check; if (cb.Tag is CategoryItem el) el.IsChecked = check; }
            }
            UpdateStatus();
        }

        // ── Search ────────────────────────────────────────────────────────

        private void LoadParamsForCategory(string categoryName)
        {
            if (string.IsNullOrWhiteSpace(categoryName))
            {
                SpAvail.Children.Clear();
                SpSel.Children.Clear();
                ApplyAvailableFilters();
                ApplySelectedFilters();
                return;
            }

            var selectedNames = SpSel.Children.OfType<Border>()
                .Select(b => (b.Tag as ParameterItem)?.Name)
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .ToHashSet(System.StringComparer.OrdinalIgnoreCase);

            var parameters = ServiceLocator.IsRevitMode
                ? ServiceLocator.RevitData.GetParameters(new List<string> { categoryName })
                : _mock.GetParameters(categoryName);

            var newSelParams = parameters.Where(p => selectedNames.Contains(p.Name)).ToList();
            var newAvailParams = parameters.Where(p => !selectedNames.Contains(p.Name)).ToList();

            SpAvail.Children.Clear();
            SpSel.Children.Clear();

            foreach (var p in newSelParams)
                SpSel.Children.Add(CreateSelectedRow(p));

            foreach (var p in newAvailParams)
                SpAvail.Children.Add(CreateAvailableRow(p));

            ApplyAvailableFilters();
            ApplySelectedFilters();
        }

        private void DiscFilter_CheckedChanged(object sender, RoutedEventArgs e)
        {
            if (!IsLoaded) return;
            if (sender is CheckBox cb)
            {
                var content = cb.Content?.ToString();
                var cbAll = CbDisc.Items.OfType<CheckBox>()
                    .FirstOrDefault(x => x.Content?.ToString() == "<All Disciplines>");
                var others = CbDisc.Items.OfType<CheckBox>()
                    .Where(x => x != cbAll)
                    .ToList();

                if (content == "<All Disciplines>")
                {
                    foreach (var item in others)
                        item.IsChecked = cb.IsChecked == true;
                }
                else
                {
                    if (cbAll != null)
                        cbAll.IsChecked = others.Any() && others.All(x => x.IsChecked == true);
                }
            }
            UpdateDiscFilterText();
            ApplyDisciplineFilter();
        }

        private void NormalizeDisciplineSelection()
        {
            var cbAll = CbDisc.Items.OfType<CheckBox>()
                .FirstOrDefault(x => x.Content?.ToString() == "<All Disciplines>");
            if (cbAll == null)
                return;

            var others = CbDisc.Items.OfType<CheckBox>().Where(x => x != cbAll).ToList();
            if (cbAll.IsChecked == true)
            {
                foreach (var item in others)
                    item.IsChecked = true;
                return;
            }

            if (others.Any() && others.All(x => x.IsChecked == true))
                cbAll.IsChecked = true;
        }

        private void UpdateDiscFilterText()
        {
            var checkedItems = CbDisc.Items.OfType<CheckBox>().Where(c => c.IsChecked == true).Select(c => c.Content?.ToString()).ToList();
            if (CbDisc != null) CbDisc.Text = string.Join(", ", checkedItems);
        }

        private void ApplyDisciplineFilter()
        {
            if (_allCats == null) return;
            var checkedItems = CbDisc.Items.OfType<CheckBox>().Where(c => c.IsChecked == true).Select(c => c.Content?.ToString()).ToHashSet(System.StringComparer.OrdinalIgnoreCase);
            var filtered = checkedItems.Contains("<All Disciplines>")
                ? _allCats
                : _allCats.Where(c => !string.IsNullOrWhiteSpace(c.Discipline) && checkedItems.Contains(c.Discipline)).ToList();

            SpCat.Children.Clear();
            foreach (var cat in filtered)
            {
                var row = new Border
                {
                    Tag = cat,
                    Height = 22,
                    Padding = new Thickness(8, 0, 8, 0),
                    Background = BgNorm,
                    Cursor = Cursors.Hand
                };
                row.Child = new TextBlock
                {
                    Text = cat.Name,
                    FontSize = 12.5,
                    Foreground = new SolidColorBrush(Color.FromRgb(0x1A, 0x1A, 0x1A)),
                    VerticalAlignment = VerticalAlignment.Center
                };
                row.MouseEnter += (s, _) => { if ((Border)s != _selCatRow) ((Border)s).Background = BgHover; };
                row.MouseLeave += (s, _) => { if ((Border)s != _selCatRow) ((Border)s).Background = BgNorm; };
                row.MouseLeftButtonDown += Cat_Click;
                SpCat.Children.Add(row);
            }
            TbSrchCat_TextChanged(null, null);
            UpdateStatus();
        }

        private void TbSrchCat_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (SpCat == null) return;

            var q = TbSrchCat.Text?.Trim() ?? string.Empty;
            foreach (Border row in SpCat.Children)
            {
                var tb = row.Child as TextBlock;
                row.Visibility = string.IsNullOrEmpty(q) || (tb?.Text.IndexOf(q, System.StringComparison.OrdinalIgnoreCase) >= 0)
                    ? Visibility.Visible : Visibility.Collapsed;
            }
        }

        private void ApplyElementFilters()
        {
            if (SpElm == null) return;
            var q = TbSrchElm.Text?.Trim() ?? string.Empty;
            var hideUnchecked = ChkHideElm.IsChecked == true;

            foreach (var row in SpElm.Children.OfType<Border>())
            {
                if (row.Tag is not CategoryItem el) continue;
                var passText = string.IsNullOrWhiteSpace(q) || el.Name.IndexOf(q, System.StringComparison.OrdinalIgnoreCase) >= 0;
                var passChecked = !hideUnchecked || el.IsChecked;
                row.Visibility = passText && passChecked ? Visibility.Visible : Visibility.Collapsed;
            }
        }

        private void TbSrchElm_TextChanged(object sender, TextChangedEventArgs e) => ApplyElementFilters();
        private void ChkHideElm_Click(object sender, RoutedEventArgs e) => ApplyElementFilters();

        private void TbSrchAvail_TextChanged(object sender, TextChangedEventArgs e)
        {
            ApplyAvailableFilters();
        }
        private void TbSrchSel_TextChanged(object sender, TextChangedEventArgs e)
        {
            ApplySelectedFilters();
        }

        // ── Filter / Move / Sort ──────────────────────────────────────────

        private void ParamFilter_CheckedChanged(object sender, RoutedEventArgs e)
        {
            EnsureAtLeastOneChecked(ChkAvailInstance, ChkAvailType, ChkAvailReadOnly);
            UpdateParamFilterText();
            ApplyAvailableFilters();
        }

        private void UpdateParamFilterText()
        {
            var parts = new List<string>();
            if (ChkAvailInstance?.IsChecked == true) parts.Add("Instance");
            if (ChkAvailType?.IsChecked == true) parts.Add("Type");
            if (ChkAvailReadOnly?.IsChecked == true) parts.Add("Read-only");
            if (CbParamFilter != null) CbParamFilter.Text = string.Join(", ", parts);
        }

        private void SelParamFilter_CheckedChanged(object sender, RoutedEventArgs e)
        {
            EnsureAtLeastOneChecked(ChkSelInstance, ChkSelType, ChkSelReadOnly);
            UpdateSelParamFilterText();
            ApplySelectedFilters();
        }

        private void UpdateSelParamFilterText()
        {
            var parts = new List<string>();
            if (ChkSelInstance?.IsChecked == true) parts.Add("Instance");
            if (ChkSelType?.IsChecked == true) parts.Add("Type");
            if (ChkSelReadOnly?.IsChecked == true) parts.Add("Read-only");
            if (CbSelParamFilter != null) CbSelParamFilter.Text = string.Join(", ", parts);
        }

        private static void EnsureAtLeastOneChecked(CheckBox a, CheckBox b, CheckBox c)
        {
            if (a.IsChecked == true || b.IsChecked == true || c.IsChecked == true)
                return;
            a.IsChecked = true;
            b.IsChecked = true;
            c.IsChecked = true;
        }

        private static bool PassKind(ParamKind kind, CheckBox inst, CheckBox type, CheckBox readOnly)
            => kind switch
            {
                ParamKind.Instance => inst.IsChecked == true,
                ParamKind.Type => type.IsChecked == true,
                ParamKind.ReadOnly => readOnly.IsChecked == true,
                _ => false
            };

        private void ApplyAvailableFilters()
        {
            if (SpAvail == null)
                return;

            var query = TbSrchAvail?.Text?.Trim() ?? string.Empty;
            foreach (var row in SpAvail.Children.OfType<Border>())
            {
                if (row.Tag is not ParameterItem p)
                    continue;
                var pass = PassKind(p.Kind, ChkAvailInstance, ChkAvailType, ChkAvailReadOnly) &&
                           (string.IsNullOrWhiteSpace(query) || p.Name.IndexOf(query, System.StringComparison.OrdinalIgnoreCase) >= 0);
                row.Visibility = pass ? Visibility.Visible : Visibility.Collapsed;
            }
            UpdateStatus();
        }

        private void ApplySelectedFilters()
        {
            if (SpSel == null)
                return;

            var query = TbSrchSel?.Text?.Trim() ?? string.Empty;
            foreach (var row in SpSel.Children.OfType<Border>())
            {
                if (row.Tag is not ParameterItem p)
                    continue;
                var passKind = PassKind(p.Kind, ChkSelInstance, ChkSelType, ChkSelReadOnly);
                var passText = string.IsNullOrWhiteSpace(query) || p.Name.IndexOf(query, System.StringComparison.OrdinalIgnoreCase) >= 0;
                row.Visibility = passKind && passText ? Visibility.Visible : Visibility.Collapsed;
            }
            UpdateStatus();
        }

        private void BtnRight_Click(object sender, RoutedEventArgs e)
        {
            var toMove = SpAvail.Children.OfType<Border>()
                .Where(b => b.Visibility == Visibility.Visible && b.Tag is ParameterItem p && p.IsHighlighted)
                .Select(b => b.Tag as ParameterItem)
                .Where(p => p != null)
                .ToList();
            foreach (var p in toMove)
                MoveToSelected(p);
        }

        private void BtnLeft_Click(object sender, RoutedEventArgs e)
        {
            var toMove = SpSel.Children.OfType<Border>()
                .Where(b => b.Visibility == Visibility.Visible && b.Tag is ParameterItem p && p.IsHighlighted)
                .Select(b => b.Tag as ParameterItem)
                .Where(p => p != null)
                .ToList();
            foreach (var p in toMove)
                MoveToAvailable(p);
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
            ApplySelectedFilters();
        }

        private void BadgeInst_Click(object sender, MouseButtonEventArgs e)
        {
            ChkAvailInstance.IsChecked = true;
            ChkAvailType.IsChecked = false;
            ChkAvailReadOnly.IsChecked = false;
            ParamFilter_CheckedChanged(sender, e);
        }

        private void BadgeType_Click(object sender, MouseButtonEventArgs e)
        {
            ChkAvailInstance.IsChecked = false;
            ChkAvailType.IsChecked = true;
            ChkAvailReadOnly.IsChecked = false;
            ParamFilter_CheckedChanged(sender, e);
        }

        private void BadgeRo_Click(object sender, MouseButtonEventArgs e)
        {
            ChkAvailInstance.IsChecked = false;
            ChkAvailType.IsChecked = false;
            ChkAvailReadOnly.IsChecked = true;
            ParamFilter_CheckedChanged(sender, e);
        }

        // ── Toolbar / Action bar ──────────────────────────────────────────

        private void BtnSectionBox_Click(object s, RoutedEventArgs e)
            => MessageBox.Show("Section Box — Phase 3.", "SheetLink", MessageBoxButton.OK, MessageBoxImage.Information);
        private void BtnExportProjectStandards_Click(object s, RoutedEventArgs e)
        {
            var opts = ShowProjectStandardsDialog();
            if (opts == null) return;
            if (opts.Google)
            {
                MessageBox.Show("Google export will be implemented later.", "THBIM", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            var selected = string.Join(", ", opts.Categories);
            if (string.IsNullOrWhiteSpace(selected)) selected = "None";
            MessageBox.Show(
                $"Project Standards export form is ready.\nSelected: {selected}\nOpen after export: {(opts.OpenAfter ? "Yes" : "No")}",
                "THBIM", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void BtnReset_Click(object sender, RoutedEventArgs e)
        {
            if (_selCatRow != null) { _selCatRow.Background = BgNorm; _selCatRow = null; }
            _selCatName = null;
            SpElm.Children.Clear();
            SpAvail.Children.Clear();
            SpSel.Children.Clear();
            ChkAvailInstance.IsChecked = true;
            ChkAvailType.IsChecked = true;
            ChkAvailReadOnly.IsChecked = true;
            UpdateParamFilterText();

            ChkSelInstance.IsChecked = true;
            ChkSelType.IsChecked = true;
            ChkSelReadOnly.IsChecked = true;
            UpdateSelParamFilterText();

            var cbAll = CbDisc.Items.OfType<CheckBox>().FirstOrDefault(x => x.Content?.ToString() == "<All Disciplines>");
            if (cbAll != null)
            {
                cbAll.IsChecked = true;
                foreach (var item in CbDisc.Items.OfType<CheckBox>().Where(x => x != cbAll)) item.IsChecked = true;
                UpdateDiscFilterText();
            }

            TbSrchCat.Clear(); TbSrchElm.Clear(); TbSrchAvail.Clear(); TbSrchSel.Clear();
            ApplyDisciplineFilter();
        }

        private void BtnPreview_Click(object sender, RoutedEventArgs e)
            => NavigateToPreview();

        private void BtnImport_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            { Filter = "Excel Files|*.xlsx;*.xls|All Files|*.*" };
            if (dlg.ShowDialog() != true) return;
            if (!ServiceLocator.IsRevitMode) { ShowImportResult(new ImportResult(), "Import Elements"); return; }

            var handler = Services.RevitEventHandler.Instance;
            if (handler == null) { MessageBox.Show("Revit event handler not initialized.", "THBIM", MessageBoxButton.OK, MessageBoxImage.Error); return; }

            var host = Window.GetWindow(this) as SheetLinkWindow;
            host?.ShowBusy("Importing from Excel...");

            ImportResult result = null;
            handler.Enqueue(
                _ =>
                {
                    result = ServiceLocator.Excel.ImportFromExcel(dlg.FileName, RevitDocumentCache.Current);
                },
                () => Dispatcher.Invoke(() =>
                {
                    host?.HideBusy();
                    if (result != null) ShowImportResult(result, "Import Elements");
                }));
        }

        private void CbImportAction_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!IsLoaded || sender is not ComboBox cb || cb.SelectedIndex <= 0)
                return;

            BtnImport_Click(sender, new RoutedEventArgs());
            cb.SelectedIndex = 0;
        }

        private async void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            var parms = SpSel.Children.OfType<Border>()
                             .Select(b => (b.Tag as ParameterItem)?.Name)
                             .Where(n => n != null).ToList();

            if (string.IsNullOrEmpty(_selCatName) || !parms.Any())
            {
                MessageBox.Show("Select a Category and at least one Parameter.", "SheetLink",
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
                FileName = $"SheetLink_Elements_{System.DateTime.Now:yyyyMMdd_HHmm}.xlsx"
            };
            if (dlg.ShowDialog() != true) return;

            var byTypeId = ChkByTypeId?.IsChecked == true;
            await RunWithBusy("Exporting...", async () =>
            {
                if (ServiceLocator.IsRevitMode)
                    ServiceLocator.Excel.ExportCategories(
                        dlg.FileName, new List<string> { _selCatName }, parms, byTypeId, opts.IncludeInstruction);
                await Task.CompletedTask;
            }, "Export complete!");

            if (opts.OpenAfter)
                OpenFile(dlg.FileName);
        }

        // ── Status ────────────────────────────────────────────────────────

        private void UpdateStatus()
        {
            int elms = SpElm.Children.OfType<Border>()
                             .Count(b => (b.Child as StackPanel)?
                             .Children.OfType<CheckBox>().FirstOrDefault()?.IsChecked == true);
            int avail = SpAvail.Children.OfType<Border>().Count(b => b.Visibility == Visibility.Visible);
            int sel = SpSel.Children.Count;
            TbStatus.Text = $"Elements selected {elms} | parameters found {avail} | parameters selected {sel}";
            BtnExport.IsEnabled = sel > 0;
            BtnPreview.IsEnabled = GetPreviewParameters().Any();
        }
    }
}
