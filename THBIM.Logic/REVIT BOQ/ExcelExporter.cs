using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

namespace THBIM.Helpers
{
    public static class ExcelExporter
    {
        private static readonly Color HeaderBg = Color.FromArgb(198, 125, 60);   // #C67D3C — match UI
        private static readonly Color HeaderFg = Color.White;
        private static readonly Color AltRowBg = Color.FromArgb(252, 243, 232);  // light peach
        private static readonly Color TotalBg = Color.FromArgb(255, 255, 204);   // pale yellow
        private static readonly Color TotalFg = Color.FromArgb(139, 69, 19);     // dark brown

        // Column definitions per data type
        private static readonly ColumnDef[] PipeColumns = new[]
        {
            new ColumnDef("SystemName", "System",  false),
            new ColumnDef("TypeName",   "Type",    false),
            new ColumnDef("Diameter",   "Size",    false),
            new ColumnDef("Length",     "Length",   true),
        };

        private static readonly ColumnDef[] FittingColumns = new[]
        {
            new ColumnDef("FamilyName", "Family",   false),
            new ColumnDef("TypeName",   "Type",     false),
            new ColumnDef("Size",       "Size",     false),
            new ColumnDef("Count",      "Count",    true),
            new ColumnDef("Category",   "Category", false),
        };

        private static readonly ColumnDef[] InsulationColumns = new[]
        {
            new ColumnDef("SystemName",          "System",    false),
            new ColumnDef("Diameter",            "Host Size", false),
            new ColumnDef("InsulationThickness", "Thick.",    false),
            new ColumnDef("Length",              "Length",     true),
        };

        /// <summary>
        /// Xuất dữ liệu ra file .xlsx với format đẹp dùng EPPlus
        /// </summary>
        public static void ExportToExcel(Dictionary<string, IEnumerable<object>> dataMap, string filePath)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage())
                {
                    foreach (var entry in dataMap)
                    {
                        string sheetName = entry.Key;
                        var dataList = entry.Value?.ToList();
                        if (dataList == null || !dataList.Any()) continue;

                        var columns = GetColumnsForSheet(sheetName);
                        var ws = package.Workbook.Worksheets.Add(sheetName.ToUpper());

                        WriteSheet(ws, dataList, columns);
                    }

                    if (package.Workbook.Worksheets.Count == 0)
                        throw new Exception("No data to export.");

                    File.WriteAllBytes(filePath, package.GetAsByteArray());
                }
            }
            catch (IOException)
            {
                throw new Exception($"File '{filePath}' is currently open.\nPlease close it and try again.");
            }
        }

        private static void WriteSheet(ExcelWorksheet ws, List<object> data, ColumnDef[] columns)
        {
            int colCount = columns.Length;

            // === HEADER ROW ===
            for (int c = 0; c < colCount; c++)
            {
                var cell = ws.Cells[1, c + 1];
                cell.Value = columns[c].Header;
                cell.Style.Font.Bold = true;
                cell.Style.Font.Color.SetColor(HeaderFg);
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(HeaderBg);
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            }

            // === DATA ROWS ===
            int row = 2;
            foreach (var item in data)
            {
                var type = item.GetType();
                for (int c = 0; c < colCount; c++)
                {
                    var prop = type.GetProperty(columns[c].PropertyName);
                    if (prop == null) continue;

                    var val = prop.GetValue(item);
                    var cell = ws.Cells[row, c + 1];

                    if (val is int intVal)
                        cell.Value = intVal;
                    else if (val is double dblVal)
                        cell.Value = dblVal;
                    else if (val is string strVal && double.TryParse(strVal, out double parsed))
                        cell.Value = parsed;
                    else
                        cell.Value = val?.ToString() ?? "";

                    if (columns[c].IsNumeric)
                    {
                        cell.Style.Numberformat.Format = "#,##0.00";
                        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    }
                }

                // Alternate row color
                if (row % 2 == 0)
                {
                    var rowRange = ws.Cells[row, 1, row, colCount];
                    rowRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rowRange.Style.Fill.BackgroundColor.SetColor(AltRowBg);
                }

                row++;
            }

            int lastDataRow = row - 1;

            // === SUBTOTAL ROW ===
            if (data.Count > 0)
            {
                var totalRow = row;
                ws.Cells[totalRow, 1].Value = "TOTAL";
                ws.Cells[totalRow, 1].Style.Font.Bold = true;

                for (int c = 0; c < colCount; c++)
                {
                    if (columns[c].IsNumeric)
                    {
                        string colLetter = GetColumnLetter(c + 1);
                        ws.Cells[totalRow, c + 1].Formula = $"SUM({colLetter}2:{colLetter}{lastDataRow})";
                        ws.Cells[totalRow, c + 1].Style.Numberformat.Format = "#,##0.00";
                        ws.Cells[totalRow, c + 1].Style.Font.Bold = true;
                    }
                }

                var totalRange = ws.Cells[totalRow, 1, totalRow, colCount];
                totalRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                totalRange.Style.Fill.BackgroundColor.SetColor(TotalBg);
                totalRange.Style.Font.Color.SetColor(TotalFg);
                totalRange.Style.Font.Bold = true;
            }

            // === BORDERS ===
            int totalRows = data.Count > 0 ? row : lastDataRow;
            var allCells = ws.Cells[1, 1, totalRows, colCount];
            allCells.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            allCells.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            allCells.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            allCells.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            allCells.Style.Border.Top.Color.SetColor(Color.FromArgb(200, 200, 200));
            allCells.Style.Border.Bottom.Color.SetColor(Color.FromArgb(200, 200, 200));
            allCells.Style.Border.Left.Color.SetColor(Color.FromArgb(200, 200, 200));
            allCells.Style.Border.Right.Color.SetColor(Color.FromArgb(200, 200, 200));

            // Header border darker
            var headerRange = ws.Cells[1, 1, 1, colCount];
            headerRange.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
            headerRange.Style.Border.Bottom.Color.SetColor(Color.FromArgb(150, 80, 30));

            // === AUTO-FIT ===
            for (int c = 1; c <= colCount; c++)
            {
                ws.Column(c).AutoFit(12, 50);
            }

            // Row height for header
            ws.Row(1).Height = 28;
            ws.Row(1).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        }

        private static ColumnDef[] GetColumnsForSheet(string sheetName)
        {
            var lower = sheetName.ToLowerInvariant();
            if (lower.Contains("pipe") || lower.Contains("duct"))
                return PipeColumns;
            if (lower.Contains("fitting"))
                return FittingColumns;
            if (lower.Contains("insulation"))
                return InsulationColumns;

            // Fallback: pipe columns
            return PipeColumns;
        }

        private static string GetColumnLetter(int colNumber)
        {
            string letter = "";
            while (colNumber > 0)
            {
                int mod = (colNumber - 1) % 26;
                letter = (char)('A' + mod) + letter;
                colNumber = (colNumber - 1) / 26;
            }
            return letter;
        }

        private class ColumnDef
        {
            public string PropertyName { get; }
            public string Header { get; }
            public bool IsNumeric { get; }

            public ColumnDef(string propertyName, string header, bool isNumeric)
            {
                PropertyName = propertyName;
                Header = header;
                IsNumeric = isNumeric;
            }
        }
    }
}
