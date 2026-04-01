using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace THBIM
{
    public static class ExcelExportManager
    {
        public static void ExportToExcel(QtoProfile profile, string outputPath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("BOQ");

                // ==========================================
                // PHẦN 1: TỰ ĐỘNG VẼ HEADER & THÔNG TIN DỰ ÁN
                // ==========================================
                if (profile.ProjectData != null)
                {
                    ws.Cells[2, 10].Value = "Revision:";
                    ws.Cells[2, 11].Value = profile.ProjectData.Revision ?? "0";
                    ws.Cells[2, 11].Style.Font.Color.SetColor(System.Drawing.Color.Red);

                    ws.Cells[3, 1].Value = "PROJECT:";
                    ws.Cells[3, 2].Value = profile.ProjectData.ProjectName;
                    ws.Cells[3, 2].Style.Font.Color.SetColor(System.Drawing.Color.Red);

                    ws.Cells[3, 10].Value = "Measured by:";
                    ws.Cells[3, 11].Value = profile.ProjectData.MeasuredBy;
                    ws.Cells[3, 11].Style.Font.Color.SetColor(System.Drawing.Color.Red);
                }

                ws.Cells[4, 1].Value = "Measurement (3D)";
                ws.Cells[4, 10].Value = "Date:";
                ws.Cells[4, 11].Value = DateTime.Now.ToString("yyyy-MM-dd");
                ws.Cells[4, 11].Style.Font.Color.SetColor(System.Drawing.Color.Red);

                string[] headers = { "Sl. No", "Elements", "Nos", "Length", "Width", "Height", "Eff. Height", "Surface Area", "Concrete Qty", "Formwork Qty", "Remarks" };
                for (int i = 0; i < headers.Length; i++)
                {
                    ws.Cells[5, i + 1].Value = headers[i];
                }

                ws.Cells[6, 4].Value = "(m)";
                ws.Cells[6, 5].Value = "(m)";
                ws.Cells[6, 6].Value = "(m)";
                ws.Cells[6, 7].Value = "(m)";
                ws.Cells[6, 8].Value = "m2";
                ws.Cells[6, 9].Value = "m3";
                ws.Cells[6, 10].Value = "m2";

                using (var range = ws.Cells[5, 1, 6, 11])
                {
                    range.Style.Font.Bold = true;
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                }

                ws.Cells[3, 1].Style.Font.Bold = true;
                ws.Cells[4, 1].Style.Font.Bold = true;

                // ==========================================
                // PHẦN 2: ĐỔ DỮ LIỆU CẤU KIỆN
                // ==========================================
                int currentRow = 7;
                int startDataRow = 7;

                var validTables = profile.Tables.Where(t => t.Rows.Count > 0).ToList();
                var groupedByLevel = validTables.GroupBy(t => string.IsNullOrEmpty(t.LevelValue) || t.LevelValue == "All" ? "Overall Project" : t.LevelValue).ToList();
                int summaryIndex = 1;

                foreach (var levelGroup in groupedByLevel)
                {
                    // ==========================================
                    // 1. TẠO HÀNG TÊN TẦNG (Merge chiếm 2 hàng dọc)
                    // ==========================================
                    // Gộp từ Cột 2 đến Cột 11, kéo dài xuống 2 hàng (currentRow đến currentRow + 1)
                    using (var levelRange = ws.Cells[currentRow, 2, currentRow + 1, 11])
                    {
                        levelRange.Merge = true;
                        levelRange.Value = levelGroup.Key; // Điền tên Tầng
                        levelRange.Style.Font.Bold = true;
                        levelRange.Style.Font.Color.SetColor(System.Drawing.Color.DarkBlue);

                        // Căn Trái (Left) và Căn Giữa theo chiều dọc (Center) cho đẹp khi ô cao lên
                        levelRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        levelRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        levelRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        levelRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(220, 230, 241));
                    }

                    // Xử lý cột 1 (Sl. No): Gộp dọc 2 hàng và tô màu cho đồng bộ với bên kia
                    using (var slNoRange = ws.Cells[currentRow, 1, currentRow + 1, 1])
                    {
                        slNoRange.Merge = true;
                        slNoRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        slNoRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(220, 230, 241));
                    }

                    // ĐÃ SỬA: Cộng thêm 2 thay vì 1 (Vì hàng Tầng đã chiếm mất 2 dòng)
                    currentRow += 2;

                    // ==========================================
                    // 2. DÒNG SUMMARY TẦNG
                    // ==========================================
                    ws.Cells[currentRow, 1].Value = summaryIndex;
                    ws.Cells[currentRow, 2].Value = $"Summary - {levelGroup.Key}";
                    ws.Cells[currentRow, 1, currentRow, 2].Style.Font.Bold = true;
                    currentRow++;
                    summaryIndex++;

                    // 3. ĐỔ DỮ LIỆU
                    foreach (var table in levelGroup)
                    {
                        var colMapping = new Dictionary<string, int>();
                        for (int i = 0; i < table.Columns.Count; i++)
                        {
                            string headingName = table.Columns[i].Heading?.Trim();
                            if (!string.IsNullOrEmpty(headingName)) colMapping[headingName] = i;
                        }

                        foreach (var row in table.Rows)
                        {
                            ws.Cells[currentRow, 2].Value = row.MainValue;

                            // ĐÃ SỬA: TÔ MÀU ĐỎ CHO GHI CHÚ
                            if (!string.IsNullOrEmpty(row.Note))
                            {
                                ws.Cells[currentRow, 11].Value = row.Note;
                                ws.Cells[currentRow, 11].Style.Font.Color.SetColor(System.Drawing.Color.Red);
                            }

                            FillCellIfMatch(ws, currentRow, 3, "Nos", colMapping, row);
                            FillCellIfMatch(ws, currentRow, 4, "Length", colMapping, row);
                            FillCellIfMatch(ws, currentRow, 5, "Width", colMapping, row);
                            FillCellIfMatch(ws, currentRow, 6, "Height", colMapping, row);
                            FillCellIfMatch(ws, currentRow, 7, "Eff. Height", colMapping, row);
                            FillCellIfMatch(ws, currentRow, 8, "Surface Area", colMapping, row);
                            FillCellIfMatch(ws, currentRow, 9, "Concrete Qty", colMapping, row);
                            FillCellIfMatch(ws, currentRow, 10, "Formwork Qty", colMapping, row);

                            currentRow++;
                        }
                    }
                }

                // ==========================================
                // PHẦN 3: TẠO HÀNG TOTAL VÀ ĐÓNG KHUNG
                // ==========================================
                int lastDataRow = currentRow - 1; // Nhớ lại dòng chứa cấu kiện cuối cùng
                currentRow++; // ĐÃ SỬA: Cách ra 1 dòng trống trước khi in hàng GRAND TOTAL

                ws.Cells[currentRow, 2].Value = "GRAND TOTAL";
                ws.Cells[currentRow, 2].Style.Font.Color.SetColor(System.Drawing.Color.Red);

                // Gắn hàm SUM (Chỉ quét tới lastDataRow, không quét dòng trống)
                if (lastDataRow >= startDataRow)
                {
                    ws.Cells[currentRow, 8].Formula = $"SUM(H{startDataRow}:H{lastDataRow})";
                    ws.Cells[currentRow, 9].Formula = $"SUM(I{startDataRow}:I{lastDataRow})";
                    ws.Cells[currentRow, 10].Formula = $"SUM(J{startDataRow}:J{lastDataRow})";
                }

                // Tô nền vàng nhạt cho hàng TOTAL
                using (var totalRange = ws.Cells[currentRow, 1, currentRow, 11])
                {
                    totalRange.Style.Font.Bold = true;
                    totalRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    totalRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 255, 204));
                }

                // ĐÓNG KHUNG TOÀN BỘ BẢNG (Kẻ lưới tất cả các ô từ dòng 5 đến dòng TOTAL)
                using (var tableRange = ws.Cells[5, 1, currentRow, 11])
                {
                    tableRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    tableRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    tableRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    tableRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                }

                // Format số thập phân
                using (var numberRange = ws.Cells[7, 3, currentRow, 10])
                {
                    numberRange.Style.Numberformat.Format = "#,##0.00";
                }

                // Dãn cột tự động
                ws.Cells[1, 1, currentRow, 11].AutoFitColumns();

                package.SaveAs(new FileInfo(outputPath));
            }
        }

        private static void FillCellIfMatch(ExcelWorksheet ws, int rowNum, int excelColIndex, string targetHeading, Dictionary<string, int> colMapping, QtoRowProfile rowProfile)
        {
            if (colMapping.TryGetValue(targetHeading, out int valIndex))
            {
                if (valIndex >= 0 && valIndex < rowProfile.Cells.Count)
                {
                    var cell = rowProfile.Cells[valIndex];
                    if (cell.IsNumeric && cell.NumericValue != 0)
                    {
                        ws.Cells[rowNum, excelColIndex].Value = cell.NumericValue;
                    }
                    else if (!string.IsNullOrEmpty(cell.StringValue) && cell.StringValue != "N/A" && cell.StringValue != "0")
                    {
                        ws.Cells[rowNum, excelColIndex].Value = cell.StringValue;
                    }
                }
            }
        }
    }
}