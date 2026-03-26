using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using ClosedXML.Excel;
using THBIM.Models;

namespace THBIM.Services
{
    public class ImportResult
    {
        public int          Updated { get; set; }
        public List<string> Errors  { get; } = new();
        public bool         Success => Errors.Count == 0;
        public string       Summary =>
            $"Updated: {Updated} parameters" +
            (Errors.Count > 0 ? $"\nErrors: {Errors.Count}" : "");
    }

    public class ExcelService
    {
        private static readonly XLColor ColYellow = XLColor.FromHtml("#FDF0CC");
        private static readonly XLColor ColGreen  = XLColor.FromHtml("#EAF3DE");
        private static readonly XLColor ColRed    = XLColor.FromHtml("#FDEAEA");
        private static readonly XLColor ColGray   = XLColor.FromHtml("#F0F0F0");
        private static readonly XLColor ColBorder = XLColor.FromHtml("#D0D0D0");

        private readonly RevitDataService _revit;
        public ExcelService(RevitDataService revit) => _revit = revit;

        // ── EXPORT CATEGORIES ─────────────────────────────────────────────

        public void ExportCategories(string filePath, List<string> categoryNames,
                                      List<string> parameterNames, bool exportByTypeId,
                                      Action<int, string> onProgress = null)
            => ExportCategories(filePath, categoryNames, parameterNames, exportByTypeId, true, onProgress);

        public void ExportCategories(string filePath, List<string> categoryNames,
                                      List<string> parameterNames, bool exportByTypeId,
                                      bool includeInstructions,
                                      Action<int, string> onProgress = null)
        {
            using var wb = new XLWorkbook();

            for (int i = 0; i < categoryNames.Count; i++)
            {
                var cat = categoryNames[i];
                onProgress?.Invoke((i + 1) * 100 / categoryNames.Count,
                    $"[{i+1}/{categoryNames.Count}] {cat}");
                var ws = wb.Worksheets.Add(Sanitize(cat));
                WriteCategorySheet(ws, cat, parameterNames, exportByTypeId);
            }
            if (includeInstructions)
                WriteInstructionsSheetV2(wb);
            wb.SaveAs(filePath);
            onProgress?.Invoke(100, "Completed   100%");
        }

        private void WriteCategorySheet(IXLWorksheet ws, string catName,
                                         List<string> paramNames, bool byTypeId)
        {
            var elements = _revit.GetElements(catName);
            var rows = byTypeId
                ? elements.GroupBy(el => el.GetTypeId()).Select(g => g.First()).ToList()
                : elements;
            var sample = rows.FirstOrDefault();

            ws.Cell(1, 1).Value = "Element ID";
            StyleCell(ws.Cell(1, 1), ParamKind.ReadOnly, true);
            ws.Cell(1, 1).Style.Protection.Locked = true;
            ws.Column(1).Width = 16;

            for (int c = 0; c < paramNames.Count; c++)
            {
                var cell = ws.Cell(1, c + 2);
                cell.Value = paramNames[c];
                var kind = sample != null ? ResolveWritableParameter(sample, paramNames[c]).Kind : ParamKind.Instance;
                StyleCell(cell, kind);
                cell.Style.Protection.Locked = true;
                ws.Column(c + 2).Width = 22;
            }
            ws.Row(1).Height = 52;

            int row = 2;
            foreach (var el in rows)
            {
                var idCell = ws.Cell(row, 1);
                idCell.Value = el.Id.GetValue();
                idCell.Style.Fill.BackgroundColor = ColGray;
                idCell.Style.Protection.Locked = true;
                ApplyDataCellBorder(idCell);

                for (int c = 0; c < paramNames.Count; c++)
                {
                    var cell = ws.Cell(row, c + 2);
                    var resolved = ResolveWritableParameter(el, paramNames[c]);
                    if (resolved.Parameter == null)
                    {
                        cell.Style.Fill.BackgroundColor = ColGray;
                        cell.Style.Protection.Locked = true;
                        ApplyDataCellBorder(cell);
                        continue;
                    }

                    cell.Value = GetParamDisplayValue(resolved.Parameter);
                    cell.Style.Fill.BackgroundColor = resolved.Kind switch
                    {
                        ParamKind.Type => ColYellow,
                        ParamKind.Instance => ColGreen,
                        _ => XLColor.FromHtml("#FFF5F5")
                    };
                    cell.Style.Protection.Locked = resolved.Kind == ParamKind.ReadOnly;
                    ApplyDataCellBorder(cell);
                }
                row++;
            }
            ws.SheetView.FreezeRows(1);
            ws.RangeUsed()?.SetAutoFilter();
            ProtectWorksheet(ws);
        }

        // ── EXPORT SPATIAL ────────────────────────────────────────────────

        public void ExportSpatial(string filePath, bool isRooms, List<long> elementIds,
                                   List<string> paramNames,
                                   Action<int, string> onProgress = null)
        {
            using var wb = new XLWorkbook();
            var ws   = wb.Worksheets.Add(isRooms ? "Rooms" : "Spaces");
            var bic    = isRooms ? BuiltInCategory.OST_Rooms : BuiltInCategory.OST_MEPSpaces;
            var idSet  = new HashSet<long>(elementIds);
            var elems  = new FilteredElementCollector(_revit.GetDocument())
                .OfCategory(bic).WhereElementIsNotElementType()
                .Where(el => idSet.Contains(el.Id.GetValue())).ToList();
            var sample = elems.FirstOrDefault();

            ws.Cell(1, 1).Value = "Element ID";
            StyleCell(ws.Cell(1, 1), ParamKind.ReadOnly, true);
            ws.Cell(1, 1).Style.Protection.Locked = true;
            ws.Column(1).Width = 14;

            for (int c = 0; c < paramNames.Count; c++)
            {
                var cell = ws.Cell(1, c + 2);
                cell.Value = paramNames[c];
                var kind = sample != null ? ResolveWritableParameter(sample, paramNames[c]).Kind : ParamKind.Instance;
                StyleCell(cell, kind);
                cell.Style.Protection.Locked = true;
                ws.Column(c + 2).Width = 20;
            }
            ws.Row(1).Height = 52;

            for (int i = 0; i < elems.Count; i++)
            {
                onProgress?.Invoke((i + 1) * 100 / elems.Count,
                    $"Collecting data [{i+1}/{elems.Count}]");
                var el  = elems[i];
                int row = i + 2;
                var idCell = ws.Cell(row, 1);
                idCell.Value = el.Id.GetValue();
                idCell.Style.Fill.BackgroundColor = ColGray;
                idCell.Style.Protection.Locked = true;
                ApplyDataCellBorder(idCell);
                for (int c = 0; c < paramNames.Count; c++)
                {
                    var cell  = ws.Cell(row, c + 2);
                    var resolved = ResolveWritableParameter(el, paramNames[c]);
                    if (resolved.Parameter == null)
                    {
                        cell.Style.Fill.BackgroundColor = ColGray;
                        cell.Style.Protection.Locked = true;
                        ApplyDataCellBorder(cell);
                        continue;
                    }

                    cell.Value = GetParamDisplayValue(resolved.Parameter);
                    cell.Style.Fill.BackgroundColor = resolved.Kind switch
                    {
                        ParamKind.Type => ColYellow,
                        ParamKind.Instance => ColGreen,
                        _ => XLColor.FromHtml("#FFF5F5")
                    };
                    cell.Style.Protection.Locked = resolved.Kind == ParamKind.ReadOnly;
                    ApplyDataCellBorder(cell);
                }
            }
            ws.SheetView.FreezeRows(1);
            ws.RangeUsed()?.SetAutoFilter();
            ProtectWorksheet(ws);
            WriteInstructionsSheetV2(wb);
            wb.SaveAs(filePath);
            onProgress?.Invoke(100, "Completed   100%");
        }

        // ── EXPORT SCHEDULES ────────────────────────────────────────────────

        public void ExportSchedules(string filePath, List<string> scheduleNames,
                                     List<string> parameterNames, bool exportByTypeId,
                                     Action<int, string> onProgress = null)
        {
            var doc = _revit.GetDocument();
            if (doc == null) return;

            var allSchedules = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .Where(v => !v.IsTemplate && !v.IsTitleblockRevisionSchedule)
                .ToList();

            var nameSet = new HashSet<string>(scheduleNames, StringComparer.OrdinalIgnoreCase);
            var matched = allSchedules.Where(v => nameSet.Contains(v.Name)).ToList();

            using var wb = new XLWorkbook();
            for (int i = 0; i < matched.Count; i++)
            {
                var vs = matched[i];
                onProgress?.Invoke((i + 1) * 100 / matched.Count, $"[{i + 1}/{matched.Count}] {vs.Name}");
                var ws = wb.Worksheets.Add(Sanitize(vs.Name));
                WriteScheduleSheet(ws, vs, parameterNames);
            }
            WriteInstructionsSheetV2(wb);
            wb.SaveAs(filePath);
            onProgress?.Invoke(100, "Completed   100%");
        }

        private void WriteScheduleSheet(IXLWorksheet ws, ViewSchedule schedule, List<string> parameterNames)
        {
            var doc = _revit.GetDocument();
            TableSectionData body;
            try { body = schedule.GetTableData()?.GetSectionData(SectionType.Body); }
            catch { body = null; }

            // Build column map from schedule fields
            var fieldOrder = schedule.Definition.GetFieldOrder();
            var colMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            int colOffset = 0;
            foreach (var fieldId in fieldOrder)
            {
                ScheduleField field;
                try { field = schedule.Definition.GetField(fieldId); }
                catch { continue; }
                if (field == null || field.IsHidden) continue;
                var name = field.GetName();
                if (string.IsNullOrWhiteSpace(name) || colMap.ContainsKey(name)) continue;
                colMap[name] = colOffset;
                colOffset++;
            }

            // Determine which parameters to export
            var paramsToExport = parameterNames != null && parameterNames.Count > 0
                ? parameterNames.Where(p => colMap.ContainsKey(p)).ToList()
                : colMap.Keys.ToList();

            if (!paramsToExport.Any())
                paramsToExport = colMap.Keys.ToList();

            // Collect element IDs from schedule (for import round-trip)
            var elementIds = new List<long>();
            try
            {
                elementIds = new FilteredElementCollector(doc, schedule.Id)
                    .WhereElementIsNotElementType()
                    .Select(el => el.Id.GetValue())
                    .ToList();
            }
            catch { }

            // Column 1 = Element ID (locked/gray), params start at column 2
            var idHeader = ws.Cell(1, 1);
            idHeader.Value = "Element ID";
            StyleCell(idHeader, ParamKind.ReadOnly, true);
            idHeader.Style.Protection.Locked = true;
            ws.Column(1).Width = 14;

            // Write header row — parameter columns
            for (int c = 0; c < paramsToExport.Count; c++)
            {
                var cell = ws.Cell(1, c + 2);
                cell.Value = paramsToExport[c];
                var kind = ResolveScheduleFieldKind(doc, schedule, paramsToExport[c]);
                StyleCell(cell, kind);
                cell.Style.Protection.Locked = true;
                ws.Column(c + 2).Width = 22;
            }
            ws.Row(1).Height = 52;

            // Cache kind per column
            var kindPerCol = new ParamKind[paramsToExport.Count];
            for (int c = 0; c < paramsToExport.Count; c++)
                kindPerCol[c] = ResolveScheduleFieldKind(doc, schedule, paramsToExport[c]);

            // Write data rows from schedule table
            if (body != null && body.NumberOfRows > 0)
            {
                for (int r = 0; r < body.NumberOfRows; r++)
                {
                    int rowNumber = body.FirstRowNumber + r;

                    // Element ID cell
                    var idCell = ws.Cell(r + 2, 1);
                    idCell.Value = r < elementIds.Count ? elementIds[r] : 0;
                    idCell.Style.Fill.BackgroundColor = ColGray;
                    idCell.Style.Protection.Locked = true;
                    ApplyDataCellBorder(idCell);

                    for (int c = 0; c < paramsToExport.Count; c++)
                    {
                        var cell = ws.Cell(r + 2, c + 2);
                        if (colMap.TryGetValue(paramsToExport[c], out var schedCol))
                        {
                            try
                            {
                                cell.Value = schedule.GetCellText(SectionType.Body,
                                    rowNumber, body.FirstColumnNumber + schedCol) ?? string.Empty;
                            }
                            catch { cell.Value = string.Empty; }
                        }

                        // Apply color matching header kind
                        var kind = kindPerCol[c];
                        cell.Style.Fill.BackgroundColor = kind switch
                        {
                            ParamKind.Type => ColYellow,
                            ParamKind.Instance => ColGreen,
                            ParamKind.ReadOnly => XLColor.FromHtml("#FFF5F5"),
                            _ => ColGreen
                        };
                        cell.Style.Protection.Locked = kind == ParamKind.ReadOnly;
                        ApplyDataCellBorder(cell);
                    }
                }
            }

            ProtectWorksheet(ws);

            ws.SheetView.FreezeRows(1);
            ws.RangeUsed()?.SetAutoFilter();
        }

        private static ParamKind ResolveScheduleFieldKind(Document doc, ViewSchedule schedule, string fieldName)
        {
            Element sample = null;
            try
            {
                sample = new FilteredElementCollector(doc, schedule.Id)
                    .WhereElementIsNotElementType().FirstOrDefault();
            }
            catch { }

            if (sample == null) return ParamKind.Instance;

            var instanceParam = sample.LookupParameter(fieldName);
            if (instanceParam != null)
                return instanceParam.IsReadOnly ? ParamKind.ReadOnly : ParamKind.Instance;

            var typeId = sample.GetTypeId();
            if (typeId != null && typeId != ElementId.InvalidElementId)
            {
                var typeParam = doc.GetElement(typeId)?.LookupParameter(fieldName);
                if (typeParam != null)
                    return typeParam.IsReadOnly ? ParamKind.ReadOnly : ParamKind.Type;
            }
            return ParamKind.ReadOnly;
        }

        // ── IMPORT ────────────────────────────────────────────────────────

        public ImportResult ImportFromExcel(string filePath, Document doc,
                                             Action<int, string> onProgress = null)
        {
            var result = new ImportResult();
            if (!File.Exists(filePath)) { result.Errors.Add("File does not exist."); return result; }

            using var wb    = new XLWorkbook(filePath);
            var sheets      = wb.Worksheets
                .Where(ws => !ws.Name.Equals("Instructions", StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (!sheets.Any())
            {
                result.Errors.Add("No data sheet found to import.");
                return result;
            }

            for (int i = 0; i < sheets.Count; i++)
            {
                onProgress?.Invoke((i + 1) * 100 / sheets.Count, $"Importing [{sheets[i].Name}]...");
                ImportSheet(sheets[i], doc, result);
            }
            onProgress?.Invoke(100, $"Completed   100%");
            return result;
        }

        private static void ImportSheet(IXLWorksheet ws, Document doc, ImportResult result)
        {
            int headerRow = FindHeaderRow(ws);
            if (headerRow < 0) { result.Errors.Add($"Sheet '{ws.Name}': No header row."); return; }

            var colMap = BuildColMap(ws, headerRow, out int idCol);
            if (idCol < 0) { result.Errors.Add($"Sheet '{ws.Name}': No 'Element ID' column."); return; }

            int lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
            using var tx = new Transaction(doc, $"SheetLink Import — {ws.Name}");
            tx.Start();
            try
            {
                for (int r = headerRow + 1; r <= lastRow; r++)
                {
                    var idVal = ws.Cell(r, idCol).GetValue<string>()?.Trim();
                    if (!int.TryParse(idVal, out int eid)) continue;
                    var el = doc.GetElement(new ElementId(eid));
                    if (el == null) { result.Errors.Add($"Row {r}: Element {eid} not found."); continue; }

                    foreach (var kvp in colMap)
                    {
                        var pName = kvp.Key;
                        var col = kvp.Value;
                        var resolved = ResolveWritableParameter(doc, el, pName);
                        if (resolved.Parameter == null || resolved.Kind == ParamKind.ReadOnly) continue;
                        try
                        {
                            if (SetParam(resolved.Parameter, ws.Cell(r, col).GetValue<string>()?.Trim() ?? ""))
                                result.Updated++;
                        }
                        catch (Exception ex)
                        {
                            result.Errors.Add($"Row {r}, '{pName}': {ex.Message}");
                        }
                    }
                }
                tx.Commit();
            }
            catch (Exception ex)
            {
                tx.RollBack();
                result.Errors.Add($"Sheet '{ws.Name}': TX error — {ex.Message}");
            }
        }

        public ImportResult ImportSpatialFromExcel(string filePath, Document doc,
                                                    bool isRooms,
                                                    Action<int, string> onProgress = null)
        {
            var result = new ImportResult();
            if (!File.Exists(filePath)) { result.Errors.Add("File does not exist."); return result; }

            using var wb  = new XLWorkbook(filePath);
            var sheetName = isRooms ? "Rooms" : "Spaces";
            var ws        = wb.Worksheets.FirstOrDefault(s =>
                s.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
            if (ws == null) { result.Errors.Add($"Sheet '{sheetName}' not found."); return result; }

            int headerRow = FindHeaderRow(ws);
            if (headerRow < 0) { result.Errors.Add("No header row."); return result; }

            var colMap = BuildColMap(ws, headerRow, out int idCol);
            int lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
            var bic     = isRooms ? BuiltInCategory.OST_Rooms : BuiltInCategory.OST_MEPSpaces;

            using var tx = new Transaction(doc, isRooms
                ? "SheetLink Import Rooms" : "SheetLink Import Spaces");
            tx.Start();
            try
            {
                for (int r = headerRow + 1; r <= lastRow; r++)
                {
                    var idVal = idCol > 0
                        ? ws.Cell(r, idCol).GetValue<string>()?.Trim()
                        : null;
                    if (!int.TryParse(idVal, out int eid)) continue;

                    var el = doc.GetElement(new ElementId(eid));
                    if (el == null) { result.Errors.Add($"Row {r}: Element {eid} not found."); continue; }

                    foreach (var kvp in colMap)
                    {
                        var pName = kvp.Key;
                        var col = kvp.Value;
                        if (pName is "Element ID" or "Phase") continue;
                        var resolved = ResolveWritableParameter(doc, el, pName);
                        if (resolved.Parameter == null || resolved.Kind == ParamKind.ReadOnly) continue;
                        try
                        {
                            if (SetParam(resolved.Parameter, ws.Cell(r, col).GetValue<string>()?.Trim() ?? ""))
                                result.Updated++;
                        }
                        catch (Exception ex)
                        {
                            result.Errors.Add($"Row {r}, '{pName}': {ex.Message}");
                        }
                    }
                }
                tx.Commit();
            }
            catch (Exception ex)
            {
                tx.RollBack();
                result.Errors.Add($"TX error: {ex.Message}");
            }
            onProgress?.Invoke(100, "Completed   100%");
            return result;
        }

        // ── HELPERS ───────────────────────────────────────────────────────

        private static void StyleCell(IXLCell cell, ParamKind kind, bool isId = false)
        {
            cell.Style.Fill.BackgroundColor = kind switch
            {
                ParamKind.ReadOnly => isId ? ColGray : ColRed,
                ParamKind.Type     => ColYellow,
                _                  => ColGreen
            };
            cell.Style.Font.Bold     = true;
            cell.Style.Font.FontSize = 11;
            cell.Style.Alignment.WrapText   = true;
            cell.Style.Alignment.Vertical   = XLAlignmentVerticalValues.Top;
            cell.Style.Border.BottomBorder  = XLBorderStyleValues.Medium;
            cell.Style.Border.BottomBorderColor = ColBorder;
            cell.Style.Border.RightBorder   = XLBorderStyleValues.Thin;
            cell.Style.Border.RightBorderColor  = ColBorder;
        }

        private (Parameter Parameter, ParamKind Kind) ResolveWritableParameter(Element element, string paramName)
            => ResolveWritableParameter(_revit.GetDocument(), element, paramName);

        private static (Parameter Parameter, ParamKind Kind) ResolveWritableParameter(Document doc, Element element, string paramName)
        {
            if (doc == null || element == null || string.IsNullOrWhiteSpace(paramName))
                return (null, ParamKind.ReadOnly);

            var instanceParam = element.LookupParameter(paramName);
            if (instanceParam != null && !instanceParam.IsReadOnly)
                return (instanceParam, ParamKind.Instance);

            Parameter typeParam = null;
            var typeId = element.GetTypeId();
            if (typeId != null && typeId != ElementId.InvalidElementId)
            {
                var type = doc.GetElement(typeId);
                typeParam = type?.LookupParameter(paramName);
                if (typeParam != null && !typeParam.IsReadOnly)
                    return (typeParam, ParamKind.Type);
            }

            if (instanceParam != null) return (instanceParam, ParamKind.ReadOnly);
            if (typeParam != null) return (typeParam, ParamKind.ReadOnly);
            return (null, ParamKind.ReadOnly);
        }

        private static string GetParamDisplayValue(Parameter param)
            => param.StorageType switch
            {
                StorageType.String => param.AsString() ?? "",
                StorageType.Double => param.AsValueString() ?? param.AsDouble().ToString(System.Globalization.CultureInfo.InvariantCulture),
                StorageType.Integer => param.AsValueString() ?? param.AsInteger().ToString(System.Globalization.CultureInfo.InvariantCulture),
                StorageType.ElementId => param.AsValueString() ?? param.AsElementId().GetValue().ToString(System.Globalization.CultureInfo.InvariantCulture),
                _ => param.AsValueString() ?? ""
            };

        private static bool SetParam(Parameter param, string value)
        {
            if (param == null || param.IsReadOnly) return false;

            value = value?.Trim() ?? string.Empty;

            switch (param.StorageType)
            {
                case StorageType.String:
                    var old = param.AsString() ?? string.Empty;
                    if (string.Equals(old, value, StringComparison.Ordinal)) return false;
                    return param.Set(value);

                case StorageType.Integer:
                    if (int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out var iv))
                    {
                        if (param.AsInteger() == iv) return false;
                        return param.Set(iv);
                    }
                    if (int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.CurrentCulture, out iv))
                    {
                        if (param.AsInteger() == iv) return false;
                        return param.Set(iv);
                    }
                    return TrySetByValueString(param, value);

                case StorageType.Double:
                    if (TrySetByValueString(param, value))
                        return true;

                    if (double.TryParse(value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var dv))
                    {
                        var curr = param.AsDouble();
                        if (Math.Abs(curr - dv) < 1e-9) return false;
                        return param.Set(dv);
                    }
                    if (double.TryParse(value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, out dv))
                    {
                        var curr = param.AsDouble();
                        if (Math.Abs(curr - dv) < 1e-9) return false;
                        return param.Set(dv);
                    }
                    return false;

                case StorageType.ElementId:
                    if (int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out var eidv))
                    {
                        var curr = param.AsElementId()?.GetValue() ?? int.MinValue;
                        if (curr == eidv) return false;
                        return param.Set(new ElementId(eidv));
                    }
                    return TrySetByValueString(param, value);

                default:
                    return false;
            }
        }

        private static bool TrySetByValueString(Parameter param, string value)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(value)) return false;
                return param.SetValueString(value);
            }
            catch
            {
                return false;
            }
        }

        private static void ApplyDataCellBorder(IXLCell cell)
        {
            cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            cell.Style.Border.BottomBorderColor = ColBorder;
            cell.Style.Border.RightBorder = XLBorderStyleValues.Thin;
            cell.Style.Border.RightBorderColor = ColBorder;
        }

        private static void ProtectWorksheet(IXLWorksheet ws)
        {
            try
            {
                ws.Protect("SheetLink");
            }
            catch
            {
                ws.Protect();
            }
        }

        private static int FindHeaderRow(IXLWorksheet ws)
        {
            int last = ws.LastRowUsed()?.RowNumber() ?? 0;
            for (int r = 1; r <= Math.Min(5, last); r++)
            {
                int lastCol = ws.Row(r).LastCellUsed()?.Address.ColumnNumber ?? 0;
                for (int c = 1; c <= lastCol; c++)
                    if (ws.Cell(r, c).GetValue<string>()?
                        .Equals("Element ID", StringComparison.OrdinalIgnoreCase) == true)
                        return r;
            }
            return -1;
        }

        private static Dictionary<string, int> BuildColMap(IXLWorksheet ws,
                                                             int row, out int idCol)
        {
            idCol = -1;
            var map     = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            int lastCol = ws.Row(row).LastCellUsed()?.Address.ColumnNumber ?? 0;
            for (int c = 1; c <= lastCol; c++)
            {
                var val = ws.Cell(row, c).GetValue<string>()?.Trim();
                if (string.IsNullOrEmpty(val)) continue;
                if (val.Equals("Element ID", StringComparison.OrdinalIgnoreCase)) idCol = c;
                else map[val] = c;
            }
            return map;
        }

        private static void WriteInstructionsSheetV2(XLWorkbook wb)
        {
            if (wb.Worksheets.Any(ws => ws.Name == "Instructions")) return;

            var ws = wb.Worksheets.Add("Instructions");
            ws.Column("A").Width = 28;
            ws.Column("B").Width = 34;
            ws.Column("C").Width = 72;
            ws.Row(2).Height = 26;
            ws.Row(3).Height = 24;
            ws.Row(4).Height = 24;
            ws.Row(5).Height = 24;

            ws.Cell("B2").Value = "Cell Fill Colour";
            ws.Cell("C2").Value = "Description";
            ws.Range("B2:C2").Style.Font.Bold = true;
            ws.Range("B2:C2").Style.Font.FontSize = 14;

            ws.Cell("B3").Style.Fill.BackgroundColor = ColYellow;
            ws.Cell("C3").Value = "Type value";
            ws.Cell("B4").Style.Fill.BackgroundColor = ColRed;
            ws.Cell("C4").Value = "Read-only value";
            ws.Cell("B5").Style.Fill.BackgroundColor = XLColor.FromHtml("#BDBDBD");
            ws.Cell("C5").Value = "Parameter does not exist for this element";
            ws.Range("B3:C5").Style.Font.FontSize = 13;

            ws.Range("B2:C5").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Range("B2:C5").Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            ws.Range("B2:C5").Style.Border.OutsideBorderColor = ColBorder;
            ws.Range("B2:C5").Style.Border.InsideBorderColor = ColBorder;

            ws.Cell("B7").Value = "Note:";
            ws.Cell("B7").Style.Font.Bold = true;
            ws.Cell("B7").Style.Font.FontSize = 15;
            ws.Cell("B8").Value = "To create new Rooms or Spaces add new rows. Provide minimum Number, Name and Phase values.";
            ws.Cell("B9").Value = "Do not fill GUID(hidden) and Element ID for new rows. Revit will auto-assign them.";
            ws.Cell("B8").Style.Font.FontSize = 13;
            ws.Cell("B9").Style.Font.FontSize = 13;

            ws.Range("B2:C9").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Range("B2:C9").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            ws.SheetView.ZoomScale = 120;
        }

        private static void WriteInstructionsSheet(XLWorkbook wb)
        {
            if (wb.Worksheets.Any(ws => ws.Name == "Instructions")) return;
            var ws = wb.Worksheets.Add("Instructions");
            ws.Column(1).Width = 80;
            var lines = new[] {
                ("SheetLink — User Guide", true),
                ("", false),
                ("HEADER COLOURS:", true),
                ("  Yellow = Type Parameter", false),
                ("  Green = Instance Parameter", false),
                ("  Light Red = Read-only (not editable)", false),
                ("", false),
                ("IMPORT RULES:", true),
                ("  • 'Element ID' column is the key — DO NOT delete", false),
                ("  • Only edit Instance cells (green)", false),
                ("  • Save the file before importing back to Revit", false),
            };
            int r = 1;
            foreach (var (text, bold) in lines)
            {
                var cell = ws.Cell(r++, 1);
                cell.Value           = text;
                cell.Style.Font.Bold = bold;
                cell.Style.Font.FontSize = 11;
            }
        }

        private static string Sanitize(string name)
        {
            var invalid = new[] { ':', '\\', '/', '?', '*', '[', ']' };
            var clean   = new string(name.Select(c => invalid.Contains(c) ? '_' : c).ToArray());
            return clean.Length > 31 ? clean[..31] : clean;
        }
    }
}
