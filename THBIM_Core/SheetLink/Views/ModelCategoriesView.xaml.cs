using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Autodesk.Revit.DB;
using THBIM.Helpers;
using THBIM.Models;
using THBIM.Services;
using Color = System.Windows.Media.Color;
using Grid = System.Windows.Controls.Grid;
using Visibility = System.Windows.Visibility;

namespace THBIM
{
    public partial class ModelCategoriesView : UserControl
    {
        private readonly MockDataService _mock = new();
        private static readonly string[] DefaultDisciplineFilters =
        {
            "Architecture",
            "Structure",
            "Mechanical",
            "Electrical",
            "Piping",
            "Infrastructure",
            "General",
            "MEP"
        };
        private List<CategoryItem> _allCategories = new();
        private bool _isolateActive;

        public ModelCategoriesView()
        {
            InitializeComponent();
            Loaded += (_, _) => LoadCategories();
        }

        private void LoadCategories()
        {
            var checkedBefore = GetCheckedCategories().ToHashSet(StringComparer.OrdinalIgnoreCase);
            var source = ServiceLocator.IsRevitMode
                ? ServiceLocator.RevitData.GetModelCategories()
                : _mock.GetModelCategories();

            source = ApplyScopeFilter(source);
            _allCategories = source
                .Where(c => c != null && !string.IsNullOrWhiteSpace(c.Name))
                .GroupBy(c => c.Name, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .OrderBy(c => c.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (var c in _allCategories)
                c.IsChecked = checkedBefore.Contains(c.Name);

            PopulateDisciplineFilter();
            ApplyDisciplineFilter();
            RefreshSelectionDependentUi();
        }

        private List<CategoryItem> ApplyScopeFilter(IEnumerable<CategoryItem> source)
        {
            var list = source?.ToList() ?? new List<CategoryItem>();
            if (!ServiceLocator.IsRevitMode || RbWholeModel?.IsChecked == true)
                return list;

            var names = GetScopedCategoryNames();
            return list.Where(c => names.Contains(c.Name)).ToList();
        }

        private HashSet<string> GetScopedCategoryNames()
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var doc = RevitDocumentCache.Current;
            if (doc == null)
                return names;

            try
            {
                if (RbActiveView?.IsChecked == true)
                {
                    var view = doc.ActiveView;
                    if (view == null)
                        return names;

                    foreach (var el in new FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType())
                    {
                        var catName = el?.Category?.Name;
                        if (!string.IsNullOrWhiteSpace(catName))
                            names.Add(catName);
                    }
                    return names;
                }

                var uiDoc = RevitDocumentCache.CurrentUi;
                if (uiDoc == null)
                    return names;

                foreach (var id in uiDoc.Selection.GetElementIds())
                {
                    var catName = doc.GetElement(id)?.Category?.Name;
                    if (!string.IsNullOrWhiteSpace(catName))
                        names.Add(catName);
                }
            }
            catch
            {
            }

            return names;
        }

        private void PopulateDisciplineFilter()
        {
            var previousChecked = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (CbDisc.Items.Count > 0)
            {
                foreach (var item in CbDisc.Items.OfType<CheckBox>())
                {
                    if (item.IsChecked == true && item.Content != null)
                        previousChecked.Add(item.Content.ToString());
                }
            }
            if (previousChecked.Count == 0) previousChecked.Add("<All Disciplines>");

            var discoveredDisciplines = _allCategories
                .Select(c => c.Discipline)
                .Where(d => !string.IsNullOrWhiteSpace(d))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(d => d, StringComparer.OrdinalIgnoreCase)
                .ToList();

            var disciplines = new List<string>();
            foreach (var d in DefaultDisciplineFilters)
            {
                if (!disciplines.Contains(d, StringComparer.OrdinalIgnoreCase))
                    disciplines.Add(d);
            }

            foreach (var d in discoveredDisciplines)
            {
                if (!disciplines.Contains(d, StringComparer.OrdinalIgnoreCase))
                    disciplines.Add(d);
            }

            CbDisc.Items.Clear();
            
            var cbAll = new CheckBox { Content = "<All Disciplines>", Margin = new Thickness(4, 2, 4, 2) };
            cbAll.Style = (Style)FindResource("Cb");
            cbAll.Click += DiscFilter_CheckedChanged;
            cbAll.IsChecked = previousChecked.Contains("<All Disciplines>");
            CbDisc.Items.Add(cbAll);

            foreach (var d in disciplines)
            {
                var cb = new CheckBox { Content = d, Margin = new Thickness(4, 2, 4, 2) };
                cb.Style = (Style)FindResource("Cb");
                cb.Click += DiscFilter_CheckedChanged;
                cb.IsChecked = previousChecked.Contains(d);
                CbDisc.Items.Add(cb);
            }

            NormalizeDisciplineSelection();
            UpdateDiscFilterText();
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

            var checkedItems = CbDisc.Items.OfType<CheckBox>()
                .Where(c => c.IsChecked == true)
                .Select(c => c.Content?.ToString())
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            if (!checkedItems.Contains("<All Disciplines>"))
            {
                foreach (var cat in _allCategories.Where(c => !string.IsNullOrWhiteSpace(c.Discipline) && !checkedItems.Contains(c.Discipline)))
                    cat.IsChecked = false;
            }

            UpdateDiscFilterText();
            ApplyDisciplineFilter();
            RefreshSelectionDependentUi();
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

        private void ApplyDisciplineFilter()
        {
            var checkedItems = CbDisc.Items.OfType<CheckBox>().Where(c => c.IsChecked == true).Select(c => c.Content?.ToString()).ToHashSet(StringComparer.OrdinalIgnoreCase);
            var filtered = checkedItems.Contains("<All Disciplines>")
                ? _allCategories
                : _allCategories.Where(c => !string.IsNullOrWhiteSpace(c.Discipline) && checkedItems.Contains(c.Discipline));

            SpCat.Children.Clear();
            foreach (var cat in filtered)
                SpCat.Children.Add(BuildCatRow(cat));
            ApplyCategoryFilters();
        }

        private Border BuildCatRow(CategoryItem cat)
        {
            var row = new Border
            {
                Tag = cat,
                Height = 22,
                Padding = new Thickness(8, 0, 8, 0),
                Background = Brushes.Transparent,
                Cursor = Cursors.Hand
            };

            var sp = new StackPanel { Orientation = Orientation.Horizontal };
            var cb = new CheckBox
            {
                IsChecked = cat.IsChecked,
                Tag = cat,
                Margin = new Thickness(0, 0, 6, 0),
                VerticalContentAlignment = VerticalAlignment.Center
            };
            cb.Checked += CatCheckBox_Changed;
            cb.Unchecked += CatCheckBox_Changed;
            sp.Children.Add(cb);
            sp.Children.Add(new TextBlock { Text = cat.Name, FontSize = 12.5, VerticalAlignment = VerticalAlignment.Center });

            row.Child = sp;
            row.MouseEnter += (s, _) => ((Border)s).Background = new SolidColorBrush(Color.FromRgb(0xF0, 0xF6, 0xFF));
            row.MouseLeave += (s, _) => ((Border)s).Background = Brushes.Transparent;
            return row;
        }

        private void CatCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            if (sender is not CheckBox cb || cb.Tag is not CategoryItem cat)
                return;
            cat.IsChecked = cb.IsChecked == true;
            RefreshSelectionDependentUi();
        }

        private void Scope_Checked(object sender, RoutedEventArgs e)
        {
            if (IsLoaded)
                LoadCategories();
        }

        private void ChkAll_Click(object sender, RoutedEventArgs e)
        {
            var check = ChkAll.IsChecked == true;
            foreach (var row in SpCat.Children.OfType<Border>().Where(r => r.Visibility == Visibility.Visible))
            {
                if (row.Tag is CategoryItem cat)
                    cat.IsChecked = check;
                if (row.Child is StackPanel sp && sp.Children.OfType<CheckBox>().FirstOrDefault() is CheckBox cb)
                    cb.IsChecked = check;
            }
            RefreshSelectionDependentUi();
        }

        private void ApplyCategoryFilters()
        {
            var q = TbSrchCat.Text?.Trim() ?? string.Empty;
            var hideUnchecked = ChkHide.IsChecked == true;
            foreach (var row in SpCat.Children.OfType<Border>())
            {
                if (row.Tag is not CategoryItem cat)
                    continue;
                var passText = string.IsNullOrWhiteSpace(q) || cat.Name.IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0;
                var passChecked = !hideUnchecked || cat.IsChecked;
                row.Visibility = passText && passChecked ? Visibility.Visible : Visibility.Collapsed;
            }
            UpdateCategoryHeaderState();
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

        private void TbSrchCat_TextChanged(object sender, TextChangedEventArgs e) { ApplyCategoryFilters(); UpdateStatus(); }
        private void TbSrchAvail_TextChanged(object sender, TextChangedEventArgs e) => ApplyAvailableFilters();
        private void TbSrchSel_TextChanged(object sender, TextChangedEventArgs e) => ApplySelectedFilters();

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

        private void UpdateDiscFilterText()
        {
            var checkedItems = CbDisc.Items.OfType<CheckBox>().Where(c => c.IsChecked == true).Select(c => c.Content?.ToString()).ToList();
            if (CbDisc != null) CbDisc.Text = string.Join(", ", checkedItems);
        }

        private static void EnsureAtLeastOneChecked(CheckBox a, CheckBox b, CheckBox c)
        {
            if (a.IsChecked == true || b.IsChecked == true || c.IsChecked == true)
                return;
            a.IsChecked = true;
            b.IsChecked = true;
            c.IsChecked = true;
        }

        private static bool PassKind(ParamKind k, CheckBox inst, CheckBox type, CheckBox ro)
            => k switch
            {
                ParamKind.Instance => inst.IsChecked == true,
                ParamKind.Type => type.IsChecked == true,
                ParamKind.ReadOnly => ro.IsChecked == true,
                _ => false
            };

        private void ApplyAvailableFilters()
        {
            var q = TbSrchAvail.Text?.Trim() ?? string.Empty;
            foreach (var row in SpAvail.Children.OfType<Border>())
            {
                if (row.Tag is not ParameterItem p)
                    continue;
                var pass = PassKind(p.Kind, ChkAvailInstance, ChkAvailType, ChkAvailReadOnly) &&
                           (string.IsNullOrWhiteSpace(q) || p.Name.IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0);
                row.Visibility = pass ? Visibility.Visible : Visibility.Collapsed;
            }
            UpdateStatus();
        }

        private void ApplySelectedFilters()
        {
            var q = TbSrchSel.Text?.Trim() ?? string.Empty;
            foreach (var row in SpSel.Children.OfType<Border>())
            {
                if (row.Tag is not ParameterItem p)
                    continue;
                var passKind = PassKind(p.Kind, ChkSelInstance, ChkSelType, ChkSelReadOnly);
                var passText = string.IsNullOrWhiteSpace(q) || p.Name.IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0;
                row.Visibility = passKind && passText ? Visibility.Visible : Visibility.Collapsed;
            }
            UpdateStatus();
        }

        private void RefreshSelectionDependentUi()
        {
            var checkedCats = GetCheckedCategories();
            LoadParamsForCategories(checkedCats);
            BtnIsolate.IsEnabled = checkedCats.Any();
            UpdateCategoryHeaderState();
            UpdateStatus();
        }

        private void LoadParamsForCategories(List<string> categoryNames)
        {
            var cats = categoryNames?
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList() ?? new List<string>();

            var selectedNames = GetSelectedParamNames().ToHashSet(StringComparer.OrdinalIgnoreCase);
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
                    .GroupBy(p => p.Name, StringComparer.OrdinalIgnoreCase)
                    .Select(g => g.First())
                    .OrderBy(p => p.Name, StringComparer.OrdinalIgnoreCase)
                    .ToList();

            foreach (var p in parameters)
            {
                if (selectedNames.Contains(p.Name))
                    SpSel.Children.Add(CreateSelectedRow(p));
                else
                    SpAvail.Children.Add(CreateAvailableRow(p));
            }

            ApplyAvailableFilters();
            ApplySelectedFilters();
        }

        private Border CreateAvailableRow(ParameterItem p) => ParamRowHelper.CreateRow(p, MoveToSelected);
        private Border CreateSelectedRow(ParameterItem p) => ParamRowHelper.CreateRow(p, MoveToAvailable);

        private void BtnRight_Click(object sender, RoutedEventArgs e)
        {
            var toMove = SpAvail.Children.OfType<Border>()
                .Where(b => b.Visibility == Visibility.Visible && b.Tag is ParameterItem p && p.IsHighlighted)
                .Select(b => b.Tag as ParameterItem)
                .Where(p => p != null)
                .ToList();
            if (!toMove.Any())
            {
                toMove = SpAvail.Children.OfType<Border>()
                    .Where(b => b.Visibility == Visibility.Visible)
                    .Select(b => b.Tag as ParameterItem)
                    .Where(p => p != null)
                    .ToList();
            }
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
            if (!toMove.Any())
            {
                toMove = SpSel.Children.OfType<Border>()
                    .Where(b => b.Visibility == Visibility.Visible)
                    .Select(b => b.Tag as ParameterItem)
                    .Where(p => p != null)
                    .ToList();
            }
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
            => a != null && b != null && string.Equals(a.Name, b.Name, StringComparison.OrdinalIgnoreCase);

        private Border GetHighlightedSelected()
            => SpSel.Children.OfType<Border>().FirstOrDefault(b => b.Tag is ParameterItem p && p.IsHighlighted);

        private void SrtTop_Click(object s, RoutedEventArgs e) { var r = GetHighlightedSelected(); if (r == null) return; SpSel.Children.Remove(r); SpSel.Children.Insert(0, r); }
        private void SrtUp_Click(object s, RoutedEventArgs e) { var r = GetHighlightedSelected(); if (r == null) return; var i = SpSel.Children.IndexOf(r); if (i <= 0) return; SpSel.Children.Remove(r); SpSel.Children.Insert(i - 1, r); }
        private void SrtDown_Click(object s, RoutedEventArgs e) { var r = GetHighlightedSelected(); if (r == null) return; var i = SpSel.Children.IndexOf(r); if (i >= SpSel.Children.Count - 1) return; SpSel.Children.Remove(r); SpSel.Children.Insert(i + 1, r); }
        private void SrtBot_Click(object s, RoutedEventArgs e) { var r = GetHighlightedSelected(); if (r == null) return; SpSel.Children.Remove(r); SpSel.Children.Add(r); }
        private void SrtReset_Click(object s, RoutedEventArgs e)
        {
            var rows = SpSel.Children.OfType<Border>().OrderBy(b => (b.Tag as ParameterItem)?.Name, StringComparer.OrdinalIgnoreCase).ToList();
            SpSel.Children.Clear();
            foreach (var row in rows) SpSel.Children.Add(row);
            ApplySelectedFilters();
        }

        private void BadgeInst_Click(object sender, MouseButtonEventArgs e) { ChkAvailInstance.IsChecked = true; ChkAvailType.IsChecked = false; ChkAvailReadOnly.IsChecked = false; ParamFilter_CheckedChanged(sender, e); }
        private void BadgeType_Click(object sender, MouseButtonEventArgs e) { ChkAvailInstance.IsChecked = false; ChkAvailType.IsChecked = true; ChkAvailReadOnly.IsChecked = false; ParamFilter_CheckedChanged(sender, e); }
        private void BadgeRo_Click(object sender, MouseButtonEventArgs e) { ChkAvailInstance.IsChecked = false; ChkAvailType.IsChecked = false; ChkAvailReadOnly.IsChecked = true; ParamFilter_CheckedChanged(sender, e); }

        private void ChkHide_Click(object sender, RoutedEventArgs e)
        {
            ApplyCategoryFilters();
            UpdateStatus();
        }

        private List<string> GetCheckedCategories()
            => _allCategories.Where(c => c.IsChecked && !string.IsNullOrWhiteSpace(c.Name))
                .Select(c => c.Name).Distinct(StringComparer.OrdinalIgnoreCase).ToList();

        private List<string> GetSelectedParamNames()
            => SpSel.Children.OfType<Border>().Select(b => (b.Tag as ParameterItem)?.Name)
                .Where(n => !string.IsNullOrWhiteSpace(n)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();

        private List<ElementId> GetScopedElementIds(List<string> categoryNames)
        {
            var ids = ServiceLocator.RevitView.GetElementIdsByCategories(categoryNames);
            if (!ids.Any() || RbWholeModel?.IsChecked == true || !ServiceLocator.IsRevitMode)
                return ids;

            var doc = RevitDocumentCache.Current;
            if (doc == null)
                return ids;

            if (RbActiveView?.IsChecked == true)
            {
                try
                {
                    var inView = new HashSet<ElementId>(new FilteredElementCollector(doc, doc.ActiveView.Id)
                        .WhereElementIsNotElementType().ToElementIds());
                    return ids.Where(inView.Contains).ToList();
                }
                catch
                {
                    return new List<ElementId>();
                }
            }

            var uiDoc = RevitDocumentCache.CurrentUi;
            if (uiDoc == null)
                return new List<ElementId>();
            var selected = new HashSet<ElementId>(uiDoc.Selection.GetElementIds());
            return ids.Where(selected.Contains).ToList();
        }

        private void BtnSectionBox_Click(object sender, RoutedEventArgs e)
        {
            if (!ServiceLocator.IsRevitMode)
            {
                MessageBox.Show("Open Revit model before using Section Box.", "THBIM", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var cats = GetCheckedCategories();
            if (!cats.Any())
            {
                MessageBox.Show("Select at least one category.", "THBIM", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var ids = GetScopedElementIds(cats);
            if (!ids.Any())
            {
                MessageBox.Show("No elements found for current scope.", "THBIM", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var opts = ShowSectionBoxDialog(ServiceLocator.RevitView.Get3DViewNames(), ServiceLocator.RevitView.GetActive3DViewName());
            if (opts == null)
                return;

            try
            {
                var view = ServiceLocator.RevitView.CreateSectionBoxWithOptions(
                    ids, opts.ViewName, opts.OffsetMm, opts.DuplicateView, opts.DuplicateViewName, opts.IsolateElements);
                SetProgress(100, $"Section box applied in '{view.Name}'.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "THBIM", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnIsolate_Click(object sender, RoutedEventArgs e)
        {
            if (!ServiceLocator.IsRevitMode)
            {
                MessageBox.Show("Open Revit model before using isolate.", "THBIM", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                if (_isolateActive)
                {
                    ServiceLocator.RevitView.ResetIsolate();
                    _isolateActive = false;
                    BtnIsolate.Content = "Isolate Selection";
                    SetProgress(100, "Isolation cleared.");
                    return;
                }

                var ids = GetScopedElementIds(GetCheckedCategories());
                if (!ids.Any())
                {
                    MessageBox.Show("No elements found for current scope.", "THBIM", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                ServiceLocator.RevitView.IsolateElements(ids);
                _isolateActive = true;
                BtnIsolate.Content = "Unisolate Selection";
                SetProgress(100, $"Isolated {ids.Count} elements.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "THBIM", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnExportProjectStandards_Click(object sender, RoutedEventArgs e)
        {
            var opts = ShowProjectStandardsDialog();
            if (opts == null)
                return;
            if (opts.Google)
            {
                MessageBox.Show("Google export will be implemented later.", "THBIM", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            var selected = string.Join(", ", opts.Categories);
            if (string.IsNullOrWhiteSpace(selected))
                selected = "None";
            MessageBox.Show(
                $"Project Standards export form is ready.\nSelected: {selected}\nOpen after export: {(opts.OpenAfter ? "Yes" : "No")}",
                "THBIM",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        private void BtnReset_Click(object sender, RoutedEventArgs e)
        {
            foreach (var cat in _allCategories)
                cat.IsChecked = false;

            if (_isolateActive && ServiceLocator.IsRevitMode)
            {
                try { ServiceLocator.RevitView.ResetIsolate(); } catch { }
            }
            _isolateActive = false;
            BtnIsolate.Content = "Isolate Selection";
            BtnIsolate.IsEnabled = false;

            ChkAll.IsChecked = false;
            ChkHide.IsChecked = false;
            ChkAvailInstance.IsChecked = true;
            ChkAvailType.IsChecked = true;
            ChkAvailReadOnly.IsChecked = true;
            UpdateParamFilterText();

            ChkSelInstance.IsChecked = true;
            ChkSelType.IsChecked = true;
            ChkSelReadOnly.IsChecked = true;
            UpdateSelParamFilterText();

            TbSrchCat.Clear();
            TbSrchAvail.Clear();
            TbSrchSel.Clear();
            
            var cbAll = CbDisc.Items.OfType<CheckBox>().FirstOrDefault(x => x.Content?.ToString() == "<All Disciplines>");
            if (cbAll != null)
            {
                cbAll.IsChecked = true;
                foreach (var item in CbDisc.Items.OfType<CheckBox>().Where(x => x != cbAll)) item.IsChecked = true;
                UpdateDiscFilterText();
            }

            ApplyDisciplineFilter();
            SpAvail.Children.Clear();
            SpSel.Children.Clear();
            UpdateStatus();
        }

        private void BtnPreview_Click(object sender, RoutedEventArgs e)
        {
            if (!GetSelectedParamNames().Any())
            {
                MessageBox.Show("Move at least one parameter to Selected Parameters before preview.", "THBIM", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            NavigateToPreview();
        }

        private void BtnImport_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog { Filter = "Excel Files|*.xlsx;*.xls|All Files|*.*" };
            if (dlg.ShowDialog() != true)
                return;

            if (!ServiceLocator.IsRevitMode)
            {
                ShowImportResult(new ImportResult(), "Import");
                return;
            }

            var handler = Services.RevitEventHandler.Instance;
            if (handler == null)
            {
                MessageBox.Show("Revit event handler not initialized.", "THBIM", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var host = Window.GetWindow(this) as SheetLinkWindow;
            host?.ShowBusy("Importing from Excel...");

            ImportResult result = null;
            handler.Enqueue(
                _ =>
                {
                    result = ServiceLocator.Excel.ImportFromExcel(
                        dlg.FileName,
                        RevitDocumentCache.Current,
                        (__, msg) => Dispatcher.Invoke(() => host?.UpdateBusyMessage(msg)));
                },
                () => Dispatcher.Invoke(() =>
                {
                    host?.HideBusy();
                    if (result != null)
                        ShowImportResult(result, "Import");
                }));
        }

        private void CbImportAction_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!IsLoaded || sender is not ComboBox cb || cb.SelectedIndex <= 0)
                return;

            var action = (cb.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? string.Empty;
            if (action.IndexOf("Excel", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                BtnImport_Click(sender, new RoutedEventArgs());
            }
            else if (action.IndexOf("Google", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                MessageBox.Show("Google import will be implemented later.", "THBIM", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show("Preview Import form will be implemented later.", "THBIM", MessageBoxButton.OK, MessageBoxImage.Information);
            }

            cb.SelectedIndex = 0;
        }

        private async void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            var cats = GetCheckedCategories();
            var parms = GetSelectedParamNames();
            if (!cats.Any())
            {
                MessageBox.Show("Select at least one category.", "THBIM", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (!parms.Any())
            {
                MessageBox.Show("Select at least one parameter.", "THBIM", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var opts = ShowExportDialog();
            if (opts == null)
                return;
            if (opts.Google)
            {
                MessageBox.Show("Google export will be implemented later.", "THBIM", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = $"SheetLink_ModelCategories_{DateTime.Now:yyyyMMdd_HHmm}.xlsx"
            };
            if (dlg.ShowDialog() != true)
                return;

            var byTypeId = ChkByTypeId?.IsChecked == true;
            await RunWithBusy($"Exporting {cats.Count} categories...", async () =>
            {
                if (ServiceLocator.IsRevitMode)
                {
                    ServiceLocator.Excel.ExportCategories(
                        dlg.FileName,
                        cats,
                        parms,
                        byTypeId,
                        opts.IncludeInstruction,
                        (_, msg) => Dispatcher.Invoke(() => UpdateBusyMessage(msg)));
                    await Task.CompletedTask;
                    return;
                }

                using var wb = new ClosedXML.Excel.XLWorkbook();
                foreach (var cat in cats)
                {
                    var ws = wb.Worksheets.Add(cat.Length > 31 ? cat[..31] : cat);
                    ws.Cell(1, 1).Value = "Element ID";
                    for (var c = 0; c < parms.Count; c++)
                        ws.Cell(1, c + 2).Value = parms[c];
                }
                wb.SaveAs(dlg.FileName);
                await Task.CompletedTask;
            }, "Export completed.");

            if (opts.OpenAfter)
                OpenFile(dlg.FileName);
        }

        private void UpdateStatus()
        {
            var cats = GetCheckedCategories().Count;
            var avail = SpAvail.Children.OfType<Border>().Count(b => b.Visibility == Visibility.Visible);
            var sel = SpSel.Children.OfType<Border>().Count();
            TbStatus.Text = $"Model categories selected {cats} | parameters found {avail} | parameters selected {sel}";
            BtnExport.IsEnabled = sel > 0;
            BtnPreview.IsEnabled = sel > 0;
        }

        private SectionBoxOptions ShowSectionBoxDialog(List<string> viewNames, string defaultViewName)
        {
            var win = CreateDialogWindow("THBIM - Isolate Selection", 430, 430);
            var root = new Grid { Margin = new Thickness(14) };
            for (var i = 0; i < 10; i++) root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

            AddLabel(root, "THBIM", 0, 26, true);
            AddLabel(root, "Select 3D View", 1, 12, false);
            var cbView = new ComboBox { Height = 30, Margin = new Thickness(0, 0, 0, 10), ItemsSource = viewNames };
            if (!string.IsNullOrWhiteSpace(defaultViewName) && viewNames.Contains(defaultViewName)) cbView.SelectedItem = defaultViewName;
            else if (viewNames.Any()) cbView.SelectedIndex = 0;
            Grid.SetRow(cbView, 2); root.Children.Add(cbView);

            AddLabel(root, "Offset buffer from selection bounding box (mm)", 3, 12, false);
            var tbOffset = new TextBox { Height = 30, Text = "0", Padding = new Thickness(6, 4, 6, 4), Margin = new Thickness(0, 0, 0, 10) };
            Grid.SetRow(tbOffset, 4); root.Children.Add(tbOffset);

            var chkDuplicate = new CheckBox { Content = "Duplicate Selected View", Margin = new Thickness(0, 0, 0, 4) };
            Grid.SetRow(chkDuplicate, 5); root.Children.Add(chkDuplicate);
            var tbDuplicate = new TextBox { Height = 30, IsEnabled = false, Text = "Duplicate View Name", Padding = new Thickness(6, 4, 6, 4), Margin = new Thickness(0, 0, 0, 10) };
            chkDuplicate.Checked += (_, _) => tbDuplicate.IsEnabled = true;
            chkDuplicate.Unchecked += (_, _) => tbDuplicate.IsEnabled = false;
            Grid.SetRow(tbDuplicate, 6); root.Children.Add(tbDuplicate);

            var chkIsolate = new CheckBox { Content = "Isolate Selected Elements", Margin = new Thickness(0, 0, 0, 14) };
            Grid.SetRow(chkIsolate, 7); root.Children.Add(chkIsolate);
            var btnApply = new Button { Content = "Apply", Width = 140, Height = 34, HorizontalAlignment = HorizontalAlignment.Right };
            Grid.SetRow(btnApply, 8); root.Children.Add(btnApply);

            SectionBoxOptions result = null;
            btnApply.Click += (_, _) =>
            {
                if (cbView.SelectedItem == null)
                {
                    MessageBox.Show("Please select a 3D view.", "THBIM", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                if (!double.TryParse(tbOffset.Text.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out var offsetMm) &&
                    !double.TryParse(tbOffset.Text.Trim(), NumberStyles.Float, CultureInfo.CurrentCulture, out offsetMm))
                {
                    MessageBox.Show("Offset must be a number.", "THBIM", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var duplicateName = chkDuplicate.IsChecked == true ? tbDuplicate.Text.Trim() : null;
                if (chkDuplicate.IsChecked == true && string.IsNullOrWhiteSpace(duplicateName))
                    duplicateName = $"{cbView.SelectedItem}_THBIM";

                result = new SectionBoxOptions
                {
                    ViewName = cbView.SelectedItem.ToString(),
                    OffsetMm = offsetMm,
                    DuplicateView = chkDuplicate.IsChecked == true,
                    DuplicateViewName = duplicateName,
                    IsolateElements = chkIsolate.IsChecked == true
                };
                win.DialogResult = true;
            };

            win.Content = root;
            return win.ShowDialog() == true ? result : null;
        }

        private ProjectStandardsOptions ShowProjectStandardsDialog()
        {
            var win = CreateDialogWindow("THBIM - Export Project Standards", 520, 430);
            var root = new Grid { Margin = new Thickness(16) };
            for (var i = 0; i < 6; i++) root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

            AddLabel(root, "THBIM", 0, 30, true);
            var grpExport = new GroupBox { Header = "Export Options", Margin = new Thickness(0, 0, 0, 10) };
            var chkOpen = new CheckBox { Content = "Open Excel File After Export", IsChecked = true, Margin = new Thickness(10, 6, 10, 8) };
            grpExport.Content = chkOpen;
            Grid.SetRow(grpExport, 1); root.Children.Add(grpExport);

            var grpChoose = new GroupBox { Header = "Choose Categories", Margin = new Thickness(0, 0, 0, 14) };
            var panel = new StackPanel { Margin = new Thickness(10, 6, 10, 8) };
            var chkProjectInfo = new CheckBox { Content = "Project Information", IsChecked = true };
            var chkObjectStyles = new CheckBox { Content = "Object Styles", IsChecked = true };
            var chkLineStyles = new CheckBox { Content = "Line Styles", IsChecked = true };
            var chkFamilies = new CheckBox { Content = "Families", IsChecked = true };
            panel.Children.Add(chkProjectInfo); panel.Children.Add(chkObjectStyles); panel.Children.Add(chkLineStyles); panel.Children.Add(chkFamilies);
            grpChoose.Content = panel;
            Grid.SetRow(grpChoose, 2); root.Children.Add(grpChoose);

            var btnRow = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Center };
            var btnExcel = new Button { Content = "Excel", Width = 130, Height = 42, Margin = new Thickness(0, 0, 10, 0) };
            var btnGoogle = new Button { Content = "Google", Width = 130, Height = 42 };
            btnRow.Children.Add(btnExcel); btnRow.Children.Add(btnGoogle);
            Grid.SetRow(btnRow, 3); root.Children.Add(btnRow);

            ProjectStandardsOptions result = null;
            void Capture(bool google)
            {
                var list = new List<string>();
                if (chkProjectInfo.IsChecked == true) list.Add("Project Information");
                if (chkObjectStyles.IsChecked == true) list.Add("Object Styles");
                if (chkLineStyles.IsChecked == true) list.Add("Line Styles");
                if (chkFamilies.IsChecked == true) list.Add("Families");
                result = new ProjectStandardsOptions { Google = google, OpenAfter = chkOpen.IsChecked == true, Categories = list };
                win.DialogResult = true;
            }
            btnExcel.Click += (_, _) => Capture(false);
            btnGoogle.Click += (_, _) => Capture(true);

            win.Content = root;
            return win.ShowDialog() == true ? result : null;
        }

        private ExportOptions ShowExportDialog()
        {
            var win = CreateDialogWindow("THBIM - Export", 420, 260);
            var root = new Grid { Margin = new Thickness(16) };
            for (var i = 0; i < 5; i++) root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

            AddLabel(root, "Export Options", 0, 18, true);
            var chkOpen = new CheckBox { Content = "Open Excel file after export", IsChecked = true, Margin = new Thickness(0, 0, 0, 6) };
            Grid.SetRow(chkOpen, 1); root.Children.Add(chkOpen);
            var chkInstr = new CheckBox { Content = "Include Instructions sheet", IsChecked = true, Margin = new Thickness(0, 0, 0, 14) };
            Grid.SetRow(chkInstr, 2); root.Children.Add(chkInstr);

            var btnRow = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            var btnExcel = new Button { Content = "Export to Excel", Width = 130, Height = 32, Margin = new Thickness(0, 0, 8, 0) };
            var btnGoogle = new Button { Content = "Export to Google", Width = 130, Height = 32 };
            btnRow.Children.Add(btnExcel); btnRow.Children.Add(btnGoogle);
            Grid.SetRow(btnRow, 4); root.Children.Add(btnRow);

            ExportOptions result = null;
            btnExcel.Click += (_, _) =>
            {
                result = new ExportOptions { Google = false, OpenAfter = chkOpen.IsChecked == true, IncludeInstruction = chkInstr.IsChecked == true };
                win.DialogResult = true;
            };
            btnGoogle.Click += (_, _) =>
            {
                result = new ExportOptions { Google = true, OpenAfter = false, IncludeInstruction = chkInstr.IsChecked == true };
                win.DialogResult = true;
            };

            win.Content = root;
            return win.ShowDialog() == true ? result : null;
        }

        private Window CreateDialogWindow(string title, int width, int height)
            => new()
            {
                Title = title,
                Width = width,
                Height = height,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = GetSheetLinkWindow(),
                ResizeMode = ResizeMode.NoResize,
                WindowStyle = WindowStyle.ToolWindow,
                Background = new SolidColorBrush(Color.FromRgb(0xF4, 0xF4, 0xF4)),
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 12
            };

        private static void AddLabel(Grid root, string text, int row, double size, bool bold)
        {
            var tb = new TextBlock
            {
                Text = text,
                FontSize = size,
                Margin = new Thickness(0, 0, 0, 10),
                Foreground = new SolidColorBrush(Color.FromRgb(0x2F, 0x2F, 0x2F)),
                FontWeight = bold ? FontWeights.Bold : FontWeights.Normal
            };
            Grid.SetRow(tb, row);
            root.Children.Add(tb);
        }

        private sealed class SectionBoxOptions
        {
            public string ViewName { get; set; }
            public double OffsetMm { get; set; }
            public bool DuplicateView { get; set; }
            public string DuplicateViewName { get; set; }
            public bool IsolateElements { get; set; }
        }

        private sealed class ProjectStandardsOptions
        {
            public bool Google { get; set; }
            public bool OpenAfter { get; set; }
            public List<string> Categories { get; set; } = new();
        }

        private sealed class ExportOptions
        {
            public bool Google { get; set; }
            public bool OpenAfter { get; set; }
            public bool IncludeInstruction { get; set; }
        }
    }
}
