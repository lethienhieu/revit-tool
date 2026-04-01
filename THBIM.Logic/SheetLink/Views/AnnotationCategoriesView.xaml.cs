// ════════════════════════════════════════════════════
// AnnotationCategoriesView.xaml.cs
// ════════════════════════════════════════════════════
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
    public partial class AnnotationCategoriesView : UserControl
    {
        private readonly MockDataService _mock = new();
        private List<CategoryItem> _allCats = new();

        public AnnotationCategoriesView() { InitializeComponent(); Loaded += (_, _) => Load(); }

        private void Load()
        {
            _allCats = ServiceLocator.IsRevitMode
                ? ServiceLocator.RevitData.GetAnnotationCategories()
                : _mock.GetAnnotationCategories();
            NormalizeDisciplineSelection();
            ApplyDisciplineFilter();
            UpdateStatus();
        }

        private void Render(IEnumerable<CategoryItem> cats)
        {
            if (SpCat == null) return;

            SpCat.Children.Clear();
            if (cats == null) return;

            foreach (var cat in cats.Where(c => c != null)) SpCat.Children.Add(BuildRow(cat));
        }

        private Border BuildRow(CategoryItem cat)
        {
            var row = new Border { Tag = cat, Height = 22, Padding = new Thickness(8, 0, 8, 0), Background = Brushes.Transparent, Cursor = Cursors.Hand };
            var sp = new StackPanel { Orientation = Orientation.Horizontal };
            var cb = new CheckBox { IsChecked = cat.IsChecked, Tag = cat, Margin = new Thickness(0, 0, 6, 0), VerticalContentAlignment = VerticalAlignment.Center };
            cb.Checked += CatCheckBox_Changed; cb.Unchecked += CatCheckBox_Changed;
            var tb = new TextBlock { Text = cat.Name, FontSize = 12.5, Foreground = new SolidColorBrush(Color.FromRgb(0x1A, 0x1A, 0x1A)), VerticalAlignment = VerticalAlignment.Center };
            sp.Children.Add(cb); sp.Children.Add(tb); row.Child = sp;
            row.MouseEnter += (s, _) => ((Border)s).Background = new SolidColorBrush(Color.FromRgb(0xF0, 0xF6, 0xFF));
            row.MouseLeave += (s, _) => ((Border)s).Background = Brushes.Transparent;
            return row;
        }
        private void Cat_Click(object sender, RoutedEventArgs e) => CatCheckBox_Changed(sender, e);

        private void ChkAll_Click(object s, RoutedEventArgs e)
        {
            bool check = ChkAll.IsChecked == true;
            foreach (var row in SpCat.Children.OfType<Border>().Where(r => r.Visibility == Visibility.Visible))
            {
                if (row.Tag is CategoryItem cat)
                    cat.IsChecked = check;
                if (row.Child is StackPanel sp && sp.Children.OfType<CheckBox>().FirstOrDefault() is CheckBox cb)
                    cb.IsChecked = check;
            }
            CatCheckBox_Changed(null, e);
        }

        private void UpdateCategoryHeaderState()
        {
            var visible = SpCat.Children.OfType<Border>().Where(r => r.Visibility == Visibility.Visible).ToList();
            if (!visible.Any())
            {
                ChkAll.IsChecked = false;
                return;
            }
            var checkedCount = visible.Count(r => (r.Tag as CategoryItem)?.IsChecked == true);
            ChkAll.IsChecked = checkedCount == 0 ? false : checkedCount == visible.Count ? true : (bool?)null;
        }

        private void ApplyCategoryFilters()
        {
            if (SpCat == null) return;
            var q = TbSrchCat.Text?.Trim() ?? string.Empty;
            var hideUnchecked = ChkHide.IsChecked == true;
            foreach (var row in SpCat.Children.OfType<Border>())
            {
                if (row.Tag is not CategoryItem cat) continue;
                var passText = string.IsNullOrWhiteSpace(q) || cat.Name.IndexOf(q, System.StringComparison.OrdinalIgnoreCase) >= 0;
                var passChecked = !hideUnchecked || cat.IsChecked;
                row.Visibility = passText && passChecked ? Visibility.Visible : Visibility.Collapsed;
            }
            UpdateCategoryHeaderState();
        }

        private void TbSrchCat_TextChanged(object s, TextChangedEventArgs e) => ApplyCategoryFilters();
        private void TbSrchAvail_TextChanged(object s, TextChangedEventArgs e) => ApplyAvailableFilters();
        private void TbSrchSel_TextChanged(object s, TextChangedEventArgs e) => ApplySelectedFilters();

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
            var checkedItems = CbDisc.Items.OfType<CheckBox>().Where(c => c.IsChecked == true).Select(c => c.Content?.ToString()).ToHashSet(System.StringComparer.OrdinalIgnoreCase);
            var filtered = checkedItems.Contains("<All Disciplines>")
                ? _allCats
                : _allCats.Where(c => !string.IsNullOrWhiteSpace(c.Discipline) && checkedItems.Contains(c.Discipline)).ToList();

            Render(filtered);
            ApplyCategoryFilters();
        }

        private void ParamFilter_CheckedChanged(object sender, RoutedEventArgs e)
        {
            EnsureAtLeastOneChecked(ChkAvailInstance, ChkAvailType, ChkAvailReadOnly);
            UpdateParamFilterText();
            ApplyAvailableFilters();
            ApplySelectedFilters();
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

        private void BtnRight_Click(object s, RoutedEventArgs e)
        {
            var toMove = SpAvail.Children.OfType<Border>()
                .Where(b => b.Visibility == Visibility.Visible && b.Tag is ParameterItem p && p.IsHighlighted)
                .Select(b => b.Tag as ParameterItem)
                .Where(p => p != null)
                .ToList();
            foreach (var p in toMove)
                MoveToSelected(p);
        }
        private void BtnLeft_Click(object s, RoutedEventArgs e)
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
            if (p == null || SpSel.Children.OfType<Border>().Any(b => SameParam(b.Tag as ParameterItem, p)))
                return;
            var row = SpAvail.Children.OfType<Border>().FirstOrDefault(b => SameParam(b.Tag as ParameterItem, p));
            if (row != null)
                SpAvail.Children.Remove(row);
            p.IsHighlighted = false;
            SpSel.Children.Add(CreateSelectedRow(p));
            ApplyAvailableFilters();
            ApplySelectedFilters();
        }

        private void MoveToAvailable(ParameterItem p)
        {
            if (p == null || SpAvail.Children.OfType<Border>().Any(b => SameParam(b.Tag as ParameterItem, p)))
                return;
            var row = SpSel.Children.OfType<Border>().FirstOrDefault(b => SameParam(b.Tag as ParameterItem, p));
            if (row != null)
                SpSel.Children.Remove(row);
            p.IsHighlighted = false;
            SpAvail.Children.Add(CreateAvailableRow(p));
            ApplyAvailableFilters();
            ApplySelectedFilters();
        }

        private static bool SameParam(ParameterItem a, ParameterItem b)
            => a != null && b != null && string.Equals(a.Name, b.Name, System.StringComparison.OrdinalIgnoreCase);

        private Border GetHlSel() => SpSel.Children.OfType<Border>().FirstOrDefault(b => b.Tag is ParameterItem p && p.IsHighlighted);
        private void SrtTop_Click(object s, RoutedEventArgs e) { var r = GetHlSel(); if (r == null) return; SpSel.Children.Remove(r); SpSel.Children.Insert(0, r); }
        private void SrtUp_Click(object s, RoutedEventArgs e) { var r = GetHlSel(); if (r == null) return; int i = SpSel.Children.IndexOf(r); if (i > 0) { SpSel.Children.Remove(r); SpSel.Children.Insert(i - 1, r); } }
        private void SrtDown_Click(object s, RoutedEventArgs e) { var r = GetHlSel(); if (r == null) return; int i = SpSel.Children.IndexOf(r); if (i < SpSel.Children.Count - 1) { SpSel.Children.Remove(r); SpSel.Children.Insert(i + 1, r); } }
        private void SrtBot_Click(object s, RoutedEventArgs e) { var r = GetHlSel(); if (r == null) return; SpSel.Children.Remove(r); SpSel.Children.Add(r); }
        private void SrtReset_Click(object s, RoutedEventArgs e)
        {
            var items = SpSel.Children.OfType<Border>().OrderBy(b => (b.Tag as ParameterItem)?.Name, System.StringComparer.OrdinalIgnoreCase).ToList();
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

        private void ChkHide_Click(object s, RoutedEventArgs e)
        {
            ApplyCategoryFilters();
        }

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

        private void BtnReset_Click(object s, RoutedEventArgs e)
        {
            foreach (var row in SpCat.Children.OfType<Border>())
            { if (row.Child is StackPanel sp && sp.Children.OfType<CheckBox>().FirstOrDefault() is CheckBox cb) cb.IsChecked = false; }
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

            TbSrchCat.Clear(); TbSrchAvail.Clear(); TbSrchSel.Clear(); 
            ApplyDisciplineFilter();
            UpdateStatus();
        }
        private void BtnPreview_Click(object s, RoutedEventArgs e) => NavigateToPreview();
        private void BtnImport_Click(object s, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog { Filter = "Excel Files|*.xlsx;*.xls|All Files|*.*" };
            if (dlg.ShowDialog() != true) return;
            if (!ServiceLocator.IsRevitMode) { ShowImportResult(new ImportResult(), "Import"); return; }

            var handler = Services.RevitEventHandler.Instance;
            if (handler == null) { MessageBox.Show("Revit event handler not initialized.", "THBIM", MessageBoxButton.OK, MessageBoxImage.Error); return; }

            var host = Window.GetWindow(this) as SheetLinkWindow;
            host?.ShowBusy("Importing from Excel...");

            ImportResult result = null;
            handler.Enqueue(
                _ =>
                {
                    result = ServiceLocator.Excel.ImportFromExcel(dlg.FileName, RevitDocumentCache.Current,
                        (__, msg) => Dispatcher.Invoke(() => host?.UpdateBusyMessage(msg)));
                },
                () => Dispatcher.Invoke(() =>
                {
                    host?.HideBusy();
                    if (result != null) ShowImportResult(result, "Import");
                }));
        }
        private void CbImportAction_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!IsLoaded || sender is not ComboBox cb || cb.SelectedIndex <= 0)
                return;

            BtnImport_Click(sender, new RoutedEventArgs());
            cb.SelectedIndex = 0;
        }
        private async void BtnExport_Click(object s, RoutedEventArgs e)
        {
            var cats = GetChecked(); var parms = GetSelectedParamNames();
            if (!cats.Any() || !parms.Any())
            {
                MessageBox.Show("Select at least one Category and one Parameter.", "SheetLink", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var opts = ShowExportDialog();
            if (opts == null) return;
            if (opts.Google)
            {
                MessageBox.Show("Google export will be implemented later.", "THBIM", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var dlg = new Microsoft.Win32.SaveFileDialog { Filter = "Excel Files|*.xlsx", FileName = $"SheetLink_Annotation_{System.DateTime.Now:yyyyMMdd_HHmm}.xlsx" };
            if (dlg.ShowDialog() != true) return;

            var byTypeId = ChkByTypeId?.IsChecked == true;
            await RunWithBusy("Exporting...", async () =>
            {
                if (ServiceLocator.IsRevitMode)
                    ServiceLocator.Excel.ExportCategories(dlg.FileName, cats, parms, byTypeId, opts.IncludeInstruction);
                await Task.CompletedTask;
            }, "Export complete!");

            if (opts.OpenAfter)
                OpenFile(dlg.FileName);
        }
        private void CatCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            if (sender is CheckBox cb && cb.Tag is CategoryItem cat)
                cat.IsChecked = cb.IsChecked == true;
            RefreshSelectionDependentUi();
        }
        private void RefreshSelectionDependentUi()
        {
            var checkedCats = GetChecked();
            LoadParamsForCategories(checkedCats);
            UpdateCategoryHeaderState();
            UpdateStatus();
        }
        private void LoadParamsForCategories(List<string> categoryNames)
        {
            var cats = categoryNames?
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .Distinct(System.StringComparer.OrdinalIgnoreCase)
                .ToList() ?? new List<string>();

            var selectedNames = GetSelectedParamNames().ToHashSet(System.StringComparer.OrdinalIgnoreCase);
            SpAvail.Children.Clear();
            SpSel.Children.Clear();

            if (!cats.Any())
            {
                ApplyAvailableFilters();
                ApplySelectedFilters();
                return;
            }

            var parameters = ServiceLocator.IsRevitMode
                ? ServiceLocator.RevitData.GetParameters(cats)
                : cats.SelectMany(c => _mock.GetParameters(c))
                    .GroupBy(p => p.Name, System.StringComparer.OrdinalIgnoreCase)
                    .Select(g => g.First())
                    .OrderBy(p => p.Name, System.StringComparer.OrdinalIgnoreCase)
                    .ToList();

            foreach (var p in parameters)
            {
                if (selectedNames.Contains(p.Name)) SpSel.Children.Add(CreateSelectedRow(p));
                else SpAvail.Children.Add(CreateAvailableRow(p));
            }
            ApplyAvailableFilters();
            ApplySelectedFilters();
        }
        private List<string> GetChecked()
            => _allCats.Where(c => c.IsChecked).Select(c => c.Name).ToList();

        private List<string> GetSelectedParamNames()
            => SpSel.Children.OfType<Border>().Select(b => (b.Tag as ParameterItem)?.Name)
                .Where(n => !string.IsNullOrWhiteSpace(n)).Distinct(System.StringComparer.OrdinalIgnoreCase).ToList();

        private Border CreateAvailableRow(ParameterItem p) => ParamRowHelper.CreateRow(p, MoveToSelected);
        private Border CreateSelectedRow(ParameterItem p) => ParamRowHelper.CreateRow(p, MoveToAvailable);

        private void UpdateStatus()
        {
            int cats = GetChecked().Count, avail = SpAvail.Children.OfType<Border>().Count(b => b.Visibility == Visibility.Visible), sel = SpSel.Children.Count;
            TbStatus.Text = $"Annotation categories selected {cats} | parameters found {avail} | parameters selected {sel}";
            BtnExport.IsEnabled = sel > 0;
            BtnPreview.IsEnabled = GetPreviewParameters().Any();
        }
    }
}
