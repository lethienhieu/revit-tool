using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using THBIM.Helpers;
using THBIM.Models;
using THBIM.Services;

namespace THBIM
{
    internal sealed class PreviewRowData
    {
        private readonly Dictionary<string, string> _values;

        public PreviewRowData(string target, IDictionary<string, string> values)
        {
            Target = target ?? string.Empty;
            _values = values != null
                ? new Dictionary<string, string>(values, StringComparer.OrdinalIgnoreCase)
                : new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        public string Target { get; }

        public string GetValue(string parameterName)
        {
            if (string.IsNullOrWhiteSpace(parameterName))
                return string.Empty;
            return _values.TryGetValue(parameterName, out var value) ? value ?? string.Empty : string.Empty;
        }
    }

    internal static class PreviewValueHelpers
    {
        internal const int MaxRows = int.MaxValue;

        internal static PreviewRowData BuildElementRow(Document doc, Element element, string target, IReadOnlyList<string> parameters)
        {
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var p in parameters ?? Array.Empty<string>())
            {
                if (string.IsNullOrWhiteSpace(p))
                    continue;
                values[p] = GetParameterPreviewValue(doc, element, p);
            }
            return new PreviewRowData(target, values);
        }

        internal static string GetElementTypeDisplayName(Document doc, Element element)
        {
            if (doc == null || element == null)
                return string.Empty;

            var typeId = element.GetTypeId();
            if (typeId == null || typeId == ElementId.InvalidElementId)
                return string.Empty;

            if (doc.GetElement(typeId) is not ElementType et)
                return string.Empty;

            var family = et.FamilyName ?? string.Empty;
            if (string.IsNullOrWhiteSpace(family))
                return et.Name ?? string.Empty;

            return $"{family} - {et.Name}";
        }

        private static string GetParameterPreviewValue(Document doc, Element element, string paramName)
        {
            if (doc == null || element == null || string.IsNullOrWhiteSpace(paramName))
                return string.Empty;

            var instanceParam = element.LookupParameter(paramName);
            if (instanceParam != null)
            {
                var text = FormatParameterValue(instanceParam);
                if (!string.IsNullOrWhiteSpace(text))
                    return text;
            }

            var typeId = element.GetTypeId();
            if (typeId == null || typeId == ElementId.InvalidElementId)
                return instanceParam != null ? FormatParameterValue(instanceParam) : string.Empty;

            var typeParam = doc.GetElement(typeId)?.LookupParameter(paramName);
            if (typeParam != null)
                return FormatParameterValue(typeParam);

            return instanceParam != null ? FormatParameterValue(instanceParam) : string.Empty;
        }

        private static string FormatParameterValue(Parameter param)
        {
            if (param == null)
                return string.Empty;

            return param.StorageType switch
            {
                StorageType.String => param.AsString() ?? param.AsValueString() ?? string.Empty,
                StorageType.Double => param.AsValueString() ?? param.AsDouble().ToString(CultureInfo.InvariantCulture),
                StorageType.Integer => param.AsValueString() ?? param.AsInteger().ToString(CultureInfo.InvariantCulture),
                StorageType.ElementId => param.AsValueString() ?? param.AsElementId().GetValue().ToString(CultureInfo.InvariantCulture),
                _ => param.AsValueString() ?? string.Empty
            };
        }
    }

    public partial class ModelCategoriesView
    {
        internal List<string> GetPreviewTargets()
            => GetCheckedCategories();

        internal List<string> GetPreviewParameters()
            => GetSelectedParamNames();

        internal List<ParameterItem> GetSelectedParameterItems()
            => SpSel?.Children.OfType<System.Windows.Controls.Border>()
                .Select(b => b.Tag as ParameterItem)
                .Where(p => p != null && !string.IsNullOrWhiteSpace(p.Name))
                .ToList() ?? new List<ParameterItem>();

        internal List<PreviewRowData> GetPreviewRows()
        {
            var targets = GetPreviewTargets();
            var parameters = GetPreviewParameters();
            if (!targets.Any() || !parameters.Any())
                return new List<PreviewRowData>();

            var doc = RevitDocumentCache.Current;
            var rows = new List<PreviewRowData>();
            var exportByTypeId = ChkByTypeId?.IsChecked == true;

            if (ServiceLocator.IsRevitMode && doc != null)
            {
                foreach (var category in targets)
                {
                    List<Element> elements;
                    try { elements = ServiceLocator.RevitData.GetElements(category); }
                    catch { continue; }

                    IEnumerable<Element> exportRows = elements ?? Enumerable.Empty<Element>();
                    if (exportByTypeId)
                        exportRows = exportRows
                            .GroupBy(el => el.GetTypeId())
                            .Select(g => g.First());

                    foreach (var el in exportRows)
                    {
                        var target = $"{category} | {el.Id.GetValue()}";
                        rows.Add(PreviewValueHelpers.BuildElementRow(doc, el, target, parameters));
                    }
                }
            }

            if (!rows.Any())
            {
                foreach (var target in targets)
                    rows.Add(new PreviewRowData(target, new Dictionary<string, string>()));
            }

            return rows;
        }

        internal void ApplyProfile(ProfileData profile)
        {
            if (profile == null) return;
            var catNames = profile.ModelCategories?.ToHashSet(StringComparer.OrdinalIgnoreCase) ?? new HashSet<string>();
            var paramNames = profile.ModelParameters?.ToHashSet(StringComparer.OrdinalIgnoreCase) ?? new HashSet<string>();

            // Check matching categories
            foreach (var cat in _allCategories)
                cat.IsChecked = catNames.Contains(cat.Name);
            ApplyDisciplineFilter();

            // Set scope
            if (RbWholeModel != null && RbActiveView != null && RbCurrentSelection != null)
            {
                switch (profile.ModelScope)
                {
                    case "Active": RbActiveView.IsChecked = true; break;
                    case "Current": RbCurrentSelection.IsChecked = true; break;
                    default: RbWholeModel.IsChecked = true; break;
                }
            }

            // Set checkboxes
            if (ChkLinked != null) ChkLinked.IsChecked = profile.ModelIncludeLinkedFiles;
            if (ChkByTypeId != null) ChkByTypeId.IsChecked = profile.ModelExportByTypeId;

            // Reload params for checked categories, then move matching to Selected
            RefreshSelectionDependentUi();
            if (paramNames.Any())
            {
                var toMove = SpAvail.Children.OfType<Border>()
                    .Select(b => b.Tag as ParameterItem)
                    .Where(p => p != null && paramNames.Contains(p.Name))
                    .ToList();
                foreach (var p in toMove) MoveToSelected(p);
            }
        }
    }

    public partial class AnnotationCategoriesView
    {
        internal List<string> GetPreviewTargets()
            => GetChecked();

        internal List<string> GetPreviewParameters()
            => GetSelectedParamNames();

        internal List<ParameterItem> GetSelectedParameterItems()
            => SpSel?.Children.OfType<System.Windows.Controls.Border>()
                .Select(b => b.Tag as ParameterItem)
                .Where(p => p != null && !string.IsNullOrWhiteSpace(p.Name))
                .ToList() ?? new List<ParameterItem>();

        internal List<PreviewRowData> GetPreviewRows()
        {
            var targets = GetPreviewTargets();
            var parameters = GetPreviewParameters();
            if (!targets.Any() || !parameters.Any())
                return new List<PreviewRowData>();

            var doc = RevitDocumentCache.Current;
            var rows = new List<PreviewRowData>();

            if (ServiceLocator.IsRevitMode && doc != null)
            {
                foreach (var category in targets)
                {
                    List<Element> elements;
                    try { elements = ServiceLocator.RevitData.GetElements(category); }
                    catch { continue; }

                    foreach (var el in elements ?? Enumerable.Empty<Element>())
                    {
                        var target = $"{category} | {el.Id.GetValue()}";
                        rows.Add(PreviewValueHelpers.BuildElementRow(doc, el, target, parameters));
                    }
                }
            }

            if (!rows.Any())
            {
                foreach (var target in targets)
                    rows.Add(new PreviewRowData(target, new Dictionary<string, string>()));
            }

            return rows;
        }

        internal void ApplyProfile(ProfileData profile)
        {
            if (profile == null) return;
            var catNames = profile.AnnotationCategories?.ToHashSet(StringComparer.OrdinalIgnoreCase) ?? new HashSet<string>();
            var paramNames = profile.AnnotationParameters?.ToHashSet(StringComparer.OrdinalIgnoreCase) ?? new HashSet<string>();

            foreach (var cat in _allCats)
                cat.IsChecked = catNames.Contains(cat.Name);
            ApplyDisciplineFilter();

            if (ChkLinked != null) ChkLinked.IsChecked = profile.AnnotationIncludeLinked;
            if (ChkByTypeId != null) ChkByTypeId.IsChecked = profile.AnnotationExportByTypeId;

            RefreshSelectionDependentUi();
            if (paramNames.Any())
            {
                var toMove = SpAvail.Children.OfType<Border>()
                    .Select(b => b.Tag as ParameterItem)
                    .Where(p => p != null && paramNames.Contains(p.Name))
                    .ToList();
                foreach (var p in toMove) MoveToSelected(p);
            }
        }
    }

    public partial class ElementsView
    {
        internal List<string> GetPreviewTargets()
        {
            var selected = SpElm?.Children
                .OfType<Border>()
                .Select(b => b.Tag as CategoryItem)
                .Where(i => i?.IsChecked == true && !string.IsNullOrWhiteSpace(i.Name))
                .Select(i => i.Name)
                .Distinct(System.StringComparer.OrdinalIgnoreCase)
                .ToList() ?? new List<string>();

            if (!selected.Any() && !string.IsNullOrWhiteSpace(_selCatName))
                selected.Add(_selCatName);

            return selected;
        }

        internal List<string> GetPreviewParameters()
            => SpSel?.Children
                .OfType<Border>()
                .Select(b => (b.Tag as ParameterItem)?.Name)
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .Distinct(System.StringComparer.OrdinalIgnoreCase)
                .ToList() ?? new List<string>();

        internal List<ParameterItem> GetSelectedParameterItems()
            => SpSel?.Children.OfType<Border>()
                .Select(b => b.Tag as ParameterItem)
                .Where(p => p != null && !string.IsNullOrWhiteSpace(p.Name))
                .ToList() ?? new List<ParameterItem>();

        internal List<PreviewRowData> GetPreviewRows()
        {
            var targets = GetPreviewTargets();
            var parameters = GetPreviewParameters();
            if (!parameters.Any())
                return new List<PreviewRowData>();

            var doc = RevitDocumentCache.Current;
            var rows = new List<PreviewRowData>();

            if (ServiceLocator.IsRevitMode && doc != null && !string.IsNullOrWhiteSpace(_selCatName))
            {
                try
                {
                    var elements = ServiceLocator.RevitData.GetElements(_selCatName);
                    foreach (var el in elements ?? Enumerable.Empty<Element>())
                    {
                        var target = el.Id.GetValue().ToString(CultureInfo.InvariantCulture);
                        rows.Add(PreviewValueHelpers.BuildElementRow(doc, el, target, parameters));
                    }
                }
                catch
                {
                }
            }

            if (!rows.Any())
            {
                foreach (var target in targets)
                {
                    rows.Add(new PreviewRowData(target, new Dictionary<string, string>()));
                }
            }

            return rows;
        }

        internal string GetProfileSelectedCategory()
            => _selCatName ?? string.Empty;

        internal List<string> GetProfileCheckedElements()
            => SpElm?.Children
                .OfType<Border>()
                .Select(b => b.Tag as CategoryItem)
                .Where(i => i?.IsChecked == true && !string.IsNullOrWhiteSpace(i.Name))
                .Select(i => i.Name)
                .Distinct(System.StringComparer.OrdinalIgnoreCase)
                .ToList() ?? new List<string>();

        internal void ApplyProfile(ProfileData profile)
        {
            if (profile == null) return;
            var paramNames = profile.ElementsParameters?.ToHashSet(StringComparer.OrdinalIgnoreCase) ?? new HashSet<string>();

            if (ChkLinked != null) ChkLinked.IsChecked = profile.ElementsIncludeLinked;
            if (ChkByTypeId != null) ChkByTypeId.IsChecked = profile.ElementsExportByTypeId;

            // Select category if specified
            if (!string.IsNullOrWhiteSpace(profile.ElementsCategory))
            {
                foreach (var row in SpCat?.Children.OfType<Border>() ?? Enumerable.Empty<Border>())
                {
                    if (row.Tag is CategoryItem cat &&
                        string.Equals(cat.Name, profile.ElementsCategory, StringComparison.OrdinalIgnoreCase))
                    {
                        Cat_Click(row, new System.Windows.Input.MouseButtonEventArgs(
                            System.Windows.Input.Mouse.PrimaryDevice, 0,
                            System.Windows.Input.MouseButton.Left) { RoutedEvent = System.Windows.Input.Mouse.MouseDownEvent });
                        break;
                    }
                }

                // Check matching elements
                var elmNames = profile.ElementsSelected?.ToHashSet(StringComparer.OrdinalIgnoreCase) ?? new HashSet<string>();
                if (elmNames.Any())
                {
                    foreach (var row in SpElm?.Children.OfType<Border>() ?? Enumerable.Empty<Border>())
                    {
                        if (row.Tag is CategoryItem el && row.Child is StackPanel sp)
                        {
                            var cb = sp.Children.OfType<CheckBox>().FirstOrDefault();
                            if (cb != null)
                            {
                                el.IsChecked = elmNames.Contains(el.Name);
                                cb.IsChecked = el.IsChecked;
                            }
                        }
                    }
                }
            }

            // Move matching params to Selected
            if (paramNames.Any())
            {
                var toMove = SpAvail?.Children.OfType<Border>()
                    .Select(b => b.Tag as ParameterItem)
                    .Where(p => p != null && paramNames.Contains(p.Name))
                    .ToList() ?? new List<ParameterItem>();
                foreach (var p in toMove) MoveToSelected(p);
            }
        }
    }

    public partial class SchedulesView
    {
        internal List<ParameterItem> GetSelectedParameterItems()
            => SpParams?.Children.OfType<Border>()
                .Select(b => b.Tag as ParameterItem)
                .Where(p => p != null && !string.IsNullOrWhiteSpace(p.Name))
                .ToList() ?? new List<ParameterItem>();

        private List<ScheduleItem> GetPreviewSourceSchedules()
        {
            if (_selectedSchedule != null)
                return new List<ScheduleItem> { _selectedSchedule };

            var checkedItems = _all?
                .Where(s => s?.IsChecked == true)
                .ToList() ?? new List<ScheduleItem>();
            if (checkedItems.Any())
                return new List<ScheduleItem> { checkedItems.First() };

            return _all?.Where(s => s != null).Take(1).ToList() ?? new List<ScheduleItem>();
        }

        internal List<string> GetPreviewTargets()
            => GetPreviewSourceSchedules()
                .Where(s => !string.IsNullOrWhiteSpace(s.Name))
                .Select(s => s.Name)
                .Distinct(System.StringComparer.OrdinalIgnoreCase)
                .ToList();

        internal List<string> GetPreviewParameters()
        {
            var fromPanel = SpParams?.Children
                .OfType<Border>()
                .Select(b => (b.Tag as ParameterItem)?.Name)
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .Distinct(System.StringComparer.OrdinalIgnoreCase)
                .ToList() ?? new List<string>();

            if (fromPanel.Any())
                return fromPanel;

            var fallback = GetPreviewSourceSchedules().FirstOrDefault();
            if (fallback == null || !ServiceLocator.IsRevitMode || string.IsNullOrWhiteSpace(fallback.ElementId))
                return fromPanel;

            try
            {
                return ServiceLocator.RevitData.GetScheduleParameters(fallback.ElementId)
                    .Select(p => p?.Name)
                    .Where(n => !string.IsNullOrWhiteSpace(n))
                    .Distinct(System.StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }
            catch
            {
                return fromPanel;
            }
        }

        internal List<PreviewRowData> GetPreviewRows()
        {
            var targets = GetPreviewTargets();
            var parameters = GetPreviewParameters();
            if (!targets.Any() || !parameters.Any())
                return new List<PreviewRowData>();

            var doc = RevitDocumentCache.Current;
            var rows = new List<PreviewRowData>();

            foreach (var schedule in GetPreviewSourceSchedules())
            {
                if (schedule == null)
                    continue;

                if (ServiceLocator.IsRevitMode && doc != null &&
                    int.TryParse(schedule.ElementId, NumberStyles.Integer, CultureInfo.InvariantCulture, out var sid))
                {
                    try
                    {
                        if (doc.GetElement(new ElementId(sid)) is ViewSchedule vs)
                        {
                            var fromTable = BuildRowsFromScheduleTable(vs, parameters, schedule.Name);
                            if (fromTable.Any())
                            {
                                rows.AddRange(fromTable);
                                continue;
                            }

                            var fromElements = BuildRowsFromScheduleElements(doc, vs, parameters);
                            if (fromElements.Any())
                            {
                                rows.AddRange(fromElements);
                                continue;
                            }
                        }
                    }
                    catch
                    {
                    }
                }

                rows.Add(new PreviewRowData(schedule.Name, new Dictionary<string, string>()));
            }

            return rows;
        }

        private static List<PreviewRowData> BuildRowsFromScheduleTable(ViewSchedule schedule, IReadOnlyList<string> parameters, string scheduleName)
        {
            var rows = new List<PreviewRowData>();
            if (schedule == null || parameters == null || parameters.Count == 0)
                return rows;

            TableSectionData body;
            try
            {
                body = schedule.GetTableData()?.GetSectionData(SectionType.Body);
            }
            catch
            {
                return rows;
            }

            if (body == null || body.NumberOfRows <= 0 || body.NumberOfColumns <= 0)
                return rows;

            var colMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var fieldOrder = schedule.Definition.GetFieldOrder();
            var colOffset = 0;
            foreach (var fieldId in fieldOrder)
            {
                if (colOffset >= body.NumberOfColumns)
                    break;

                var field = schedule.Definition.GetField(fieldId);
                if (field == null)
                    continue;

                if (field.IsHidden)
                    continue;

                var name = field.GetName();
                if (string.IsNullOrWhiteSpace(name) || colMap.ContainsKey(name))
                    continue;

                colMap[name] = body.FirstColumnNumber + colOffset;
                colOffset++;
            }

            var totalRows = Math.Min(body.NumberOfRows, PreviewValueHelpers.MaxRows);
            for (var i = 0; i < totalRows; i++)
            {
                var rowNumber = body.FirstRowNumber + i;
                var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                var hasAny = false;

                foreach (var parameter in parameters)
                {
                    if (string.IsNullOrWhiteSpace(parameter))
                        continue;

                    string text = string.Empty;
                    if (colMap.TryGetValue(parameter, out var colNumber))
                    {
                        try
                        {
                            text = schedule.GetCellText(SectionType.Body, rowNumber, colNumber) ?? string.Empty;
                        }
                        catch
                        {
                            text = string.Empty;
                        }
                    }

                    if (!string.IsNullOrWhiteSpace(text))
                        hasAny = true;

                    values[parameter] = text;
                }

                if (!hasAny)
                    continue;

                rows.Add(new PreviewRowData($"{scheduleName} - Row {i + 1}", values));
            }

            return rows;
        }

        private static List<PreviewRowData> BuildRowsFromScheduleElements(Document doc, ViewSchedule schedule, IReadOnlyList<string> parameters)
        {
            var rows = new List<PreviewRowData>();
            if (doc == null || schedule == null || parameters == null || parameters.Count == 0)
                return rows;

            try
            {
                foreach (var el in new FilteredElementCollector(doc, schedule.Id)
                    .WhereElementIsNotElementType()
                    .Take(PreviewValueHelpers.MaxRows))
                {
                    var target = el?.Id?.GetValue().ToString(CultureInfo.InvariantCulture) ?? schedule.Name;
                    rows.Add(PreviewValueHelpers.BuildElementRow(doc, el, target, parameters));
                }
            }
            catch
            {
                return new List<PreviewRowData>();
            }

            return rows;
        }

        internal List<string> GetProfileCheckedSchedules()
            => _all?
                .Where(s => s?.IsChecked == true && !string.IsNullOrWhiteSpace(s.Name))
                .Select(s => s.Name)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList() ?? new List<string>();

        internal bool GetProfileExportByTypeId()
            => ChkExportByTypeId?.IsChecked == true;

        internal void ApplyProfile(ProfileData profile)
        {
            if (profile == null) return;
            var schedNames = profile.Schedules?.ToHashSet(StringComparer.OrdinalIgnoreCase) ?? new HashSet<string>();

            if (ChkExportByTypeId != null) ChkExportByTypeId.IsChecked = profile.ScheduleExportByTypeId;

            // Check matching schedules
            ScheduleItem firstChecked = null;
            foreach (var s in _all)
            {
                s.IsChecked = schedNames.Contains(s.Name);
                if (s.IsChecked && firstChecked == null) firstChecked = s;
            }

            RenderList(_all);

            // Select first checked schedule to load its parameters
            if (firstChecked != null)
                SelectSchedule(firstChecked);

            UpdateStatus();
        }
    }

    public partial class SpatialView
    {
        internal List<ParameterItem> GetSelectedParameterItems()
            => SpSel?.Children.OfType<Border>()
                .Select(b => b.Tag as ParameterItem)
                .Where(p => p != null && !string.IsNullOrWhiteSpace(p.Name))
                .ToList() ?? new List<ParameterItem>();

        internal List<string> GetPreviewTargets()
        {
            var checkedItems = _currentItems?
                .Where(i => i?.IsChecked == true)
                .ToList() ?? new List<SpatialItem>();

            var source = checkedItems.Any()
                ? checkedItems
                : _currentItems ?? new List<SpatialItem>();

            return source
                .Select(i => i?.ElementId > 0
                    ? i.ElementId.ToString(CultureInfo.InvariantCulture)
                    : string.IsNullOrWhiteSpace(i?.Name) ? i?.Number : $"{i.Number} - {i.Name}")
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        internal List<string> GetPreviewParameters()
            => SpSel?.Children
                .OfType<Border>()
                .Select(b => (b.Tag as ParameterItem)?.Name)
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .Distinct(System.StringComparer.OrdinalIgnoreCase)
                .ToList() ?? new List<string>();

        internal List<PreviewRowData> GetPreviewRows()
        {
            var parameters = GetPreviewParameters();
            if (!parameters.Any())
                return new List<PreviewRowData>();

            var checkedItems = _currentItems?
                .Where(i => i?.IsChecked == true)
                .Take(PreviewValueHelpers.MaxRows)
                .ToList() ?? new List<SpatialItem>();

            var selected = checkedItems.Any()
                ? checkedItems
                : _currentItems?
                    .Take(PreviewValueHelpers.MaxRows)
                    .ToList() ?? new List<SpatialItem>();

            if (!selected.Any())
                return new List<PreviewRowData>();

            var doc = RevitDocumentCache.Current;
            var rows = new List<PreviewRowData>();

            foreach (var item in selected)
            {
                var target = item.ElementId > 0
                    ? item.ElementId.ToString(CultureInfo.InvariantCulture)
                    : string.IsNullOrWhiteSpace(item.Name) ? item.Number : $"{item.Number} - {item.Name}";

                Element sample = null;
                if (ServiceLocator.IsRevitMode && doc != null && item.ElementId > 0)
                {
                    try { sample = doc.GetElement(ElementIdCompat.CreateId(item.ElementId)); }
                    catch { sample = null; }
                }

                rows.Add(PreviewValueHelpers.BuildElementRow(doc, sample, target, parameters));
            }

            return rows;
        }

        internal string GetPreviewSourceLabel()
            => _isRooms ? "Rooms" : "Spaces";

        internal string GetProfileSpatialType()
            => _isRooms ? "Rooms" : "Spaces";

        internal List<long> GetProfileSelectedElementIds()
            => _currentItems?
                .Where(i => i?.IsChecked == true)
                .Select(i => i.ElementId)
                .Where(id => id > 0)
                .Distinct()
                .ToList() ?? new List<long>();

        internal bool GetProfileIncludeLinked()
            => ChkLinked?.IsChecked == true;

        internal void ApplyProfile(ProfileData profile)
        {
            if (profile == null) return;
            var selectedIds = profile.SpatialSelected?.ToHashSet() ?? new HashSet<long>();
            var paramNames = profile.SpatialParameters?.ToHashSet(StringComparer.OrdinalIgnoreCase) ?? new HashSet<string>();

            // Switch spatial type if needed
            bool wantRooms = string.Equals(profile.SpatialType, "Rooms", StringComparison.OrdinalIgnoreCase);
            if (wantRooms != _isRooms)
            {
                _isRooms = wantRooms;
                LoadData();
            }

            if (ChkLinked != null) ChkLinked.IsChecked = profile.SpatialIncludeLinked;

            // Check matching spatial items
            if (selectedIds.Any())
            {
                foreach (var item in _currentItems)
                    item.IsChecked = selectedIds.Contains(item.ElementId);

                // Update UI checkboxes
                foreach (var row in SpSpatial?.Children.OfType<Border>() ?? Enumerable.Empty<Border>())
                {
                    if (row.Tag is SpatialItem si && row.Child is System.Windows.Controls.Grid g)
                    {
                        var cb = g.Children.OfType<CheckBox>().FirstOrDefault();
                        if (cb != null) cb.IsChecked = si.IsChecked;
                    }
                }
            }

            // Move matching params to Selected
            if (paramNames.Any())
            {
                var toMove = SpAvail?.Children.OfType<Border>()
                    .Select(b => b.Tag as ParameterItem)
                    .Where(p => p != null && paramNames.Contains(p.Name))
                    .ToList() ?? new List<ParameterItem>();
                foreach (var p in toMove) MoveToSelected(p);
            }

            UpdateStatus();
        }
    }
}
