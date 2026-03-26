using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows;

namespace THBIM.Helpers
{
    public static class ExcelExporter
    {
        /// <summary>
        /// Xuất dữ liệu ra file CSV (Excel mở được ngay) mà KHÔNG cần thư viện ngoài
        /// </summary>
        public static void ExportToCsv(Dictionary<string, IEnumerable<object>> dataMap, string filePath)
        {
            try
            {
                var sb = new StringBuilder();

                // Do CSV không chia Sheet được như Excel, ta sẽ ghi nối tiếp nhau
                // và ngăn cách bằng tên Sheet
                foreach (var entry in dataMap)
                {
                    string sheetName = entry.Key;
                    var dataList = entry.Value?.ToList();

                    if (dataList == null || !dataList.Any()) continue;

                    // 1. Ghi tên Sheet làm tiêu đề lớn
                    sb.AppendLine($"--- {sheetName.ToUpper()} ---");

                    // 2. Lấy Header (Tên cột)
                    var firstItem = dataList.First();
                    var properties = firstItem.GetType().GetProperties();

                    // Ghi dòng Header
                    var headerLine = string.Join(",", properties.Select(p => EscapeCsv(p.Name)));
                    sb.AppendLine(headerLine);

                    // 3. Ghi dữ liệu từng dòng
                    foreach (var item in dataList)
                    {
                        var values = new List<string>();
                        foreach (var prop in properties)
                        {
                            var val = prop.GetValue(item);

                            // Format ngày tháng hoặc số liệu nếu cần
                            if (val is DateTime d)
                                values.Add(EscapeCsv(d.ToString("dd/MM/yyyy")));
                            else if (val is double || val is decimal || val is float)
                                values.Add(val.ToString()); // Giữ nguyên số để Excel tính toán
                            else
                                values.Add(EscapeCsv(val?.ToString() ?? ""));
                        }
                        sb.AppendLine(string.Join(",", values));
                    }

                    // Thêm 2 dòng trống để ngăn cách giữa các bảng
                    sb.AppendLine();
                    sb.AppendLine();
                }

                // 4. Lưu file (Dùng Encoding.UTF8 để hỗ trợ Tiếng Việt)
                // Lưu ý: Excel đôi khi cần BOM để hiểu UTF8, nên ta dùng Encoding.UTF8 với BOM
                File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
            }
            catch (IOException)
            {
                throw new Exception($"File '{filePath}' is currently open.\nPlease try again.");
            }
            catch (Exception ex)
            {
                throw new Exception("Error writing CSV file: " + ex.Message);
            }
        }

        // Hàm xử lý dấu phẩy trong nội dung (để không bị nhảy cột sai)
        private static string EscapeCsv(string input)
        {
            if (string.IsNullOrEmpty(input)) return "";

            // Nếu nội dung có chứa dấu phẩy, xuống dòng hoặc dấu ngoặc kép -> Cần bao quanh bằng dấu ngoặc kép
            if (input.Contains(",") || input.Contains("\"") || input.Contains("\n") || input.Contains("\r"))
            {
                // Thay thế dấu " bằng "" (quy tắc CSV)
                return "\"" + input.Replace("\"", "\"\"") + "\"";
            }
            return input;
        }
    }
}