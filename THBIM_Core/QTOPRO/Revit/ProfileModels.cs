using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace THBIM
{
    // ==========================================
    // 1. THÔNG TIN DỰ ÁN DÙNG CHO EXCEL HEADER
    // ==========================================
    public class ProjectInfo
    {
        public string ProjectName { get; set; }
        public string MeasuredBy { get; set; }
        public string Revision { get; set; }
    }

    // ==========================================
    // 2. CÁC CLASS DỮ LIỆU ĐỂ LƯU THÀNH CHỮ (JSON)
    // ==========================================
    public class QtoProfile
    {
        public string ProfileName { get; set; }
        public string FilePath { get; set; }
        public ProjectInfo ProjectData { get; set; } = new ProjectInfo(); // Lưu thông tin dự án
        public List<QtoTableProfile> Tables { get; set; } = new List<QtoTableProfile>();
        public Dictionary<string, double> HistorySnapshot { get; set; } = new Dictionary<string, double>();
    }

    public class QtoTableProfile
    {
        public string CategoryName { get; set; }
        public string LevelParameter { get; set; }
        public string LevelValue { get; set; }
        public string MainParameter { get; set; }
        public string MainHeading { get; set; }
        public List<QtoColumnProfile> Columns { get; set; } = new List<QtoColumnProfile>();
        public List<QtoRowProfile> Rows { get; set; } = new List<QtoRowProfile>();
    }

    public class QtoColumnProfile
    {
        public string ColumnId { get; set; }
        public string ParameterName { get; set; }
        public string Heading { get; set; }
    }

    public class QtoRowProfile
    {
        public string MainValue { get; set; }
        public string Note { get; set; }
        // MỚI: Thêm list này để "Chụp" lại giá trị từng ô vuông (chữ/số) trên bảng
        public List<QtoCellProfile> Cells { get; set; } = new List<QtoCellProfile>();
    }

    // MỚI: Class lưu trữ thông tin của 1 ô dữ liệu
    public class QtoCellProfile
    {
        public string StringValue { get; set; }
        public double NumericValue { get; set; }
        public bool IsNumeric { get; set; }
    }

    // ==========================================
    // 3. BỘ CÔNG CỤ XỬ LÝ ĐỌC/GHI FILE .TXT VÀ TRÍ NHỚ (REGISTRY)
    // ==========================================
    public static class ProfileManager
    {
        private static string AppDataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "THBIM_QTOPRO");
        private static string RegistryFile = Path.Combine(AppDataFolder, "ProfileRegistry.json");

        public static void SaveRegistry(Dictionary<string, string> paths)
        {
            if (!Directory.Exists(AppDataFolder)) Directory.CreateDirectory(AppDataFolder);
            string json = JsonSerializer.Serialize(paths, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(RegistryFile, json);
        }

        public static Dictionary<string, string> LoadRegistry()
        {
            if (!File.Exists(RegistryFile)) return new Dictionary<string, string>();
            try
            {
                string json = File.ReadAllText(RegistryFile);
                return JsonSerializer.Deserialize<Dictionary<string, string>>(json) ?? new Dictionary<string, string>();
            }
            catch { return new Dictionary<string, string>(); }
        }

        public static void SaveToFile(QtoProfile profile, string filePath)
        {
            // ========================================================
            // BƯỚC 2: CHỤP ẢNH TẤT CẢ CÁC THAM SỐ SỐ (BỎ QUA HEADING)
            // ========================================================
            if (profile.HistorySnapshot == null) profile.HistorySnapshot = new Dictionary<string, double>();
            profile.HistorySnapshot.Clear();

            foreach (var table in profile.Tables)
            {
                for (int r = 0; r < table.Rows.Count; r++)
                {
                    var row = table.Rows[r];
                    for (int c = 0; c < table.Columns.Count; c++)
                    {
                        var col = table.Columns[c];
                        // Nếu ô đó chứa giá trị Số -> Bắt đầu chụp ảnh
                        if (c < row.Cells.Count && row.Cells[c].IsNumeric)
                        {
                            // Chìa khóa vàng: Category + Level + Tên cấu kiện + TÊN PARAMETER GỐC
                            string uniqueKey = $"{table.CategoryName}_{table.LevelValue}_{row.MainValue}_{col.ParameterName}";
                            profile.HistorySnapshot[uniqueKey] = row.Cells[c].NumericValue;
                        }
                    }
                }
            }
            // ========================================================

            var options = new System.Text.Json.JsonSerializerOptions { WriteIndented = true };
            string jsonString = System.Text.Json.JsonSerializer.Serialize(profile, options);
            System.IO.File.WriteAllText(filePath, jsonString);
        }

        public static QtoProfile LoadFromFile(string filePath)
        {
            if (!File.Exists(filePath)) return null;
            try
            {
                string jsonString = File.ReadAllText(filePath);
                var profile = JsonSerializer.Deserialize<QtoProfile>(jsonString);
                if (profile != null) profile.FilePath = filePath;
                return profile;
            }
            catch { return null; }
        }
    }
}