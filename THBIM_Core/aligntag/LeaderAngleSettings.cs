using System;
using System.IO;
using System.Text.Json;

namespace THBIM
{
    public static class LeaderAngleSettings
    {
        private static readonly string SettingsDir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "THBIM", "AlignTag");

        private static readonly string SettingsPath =
            Path.Combine(SettingsDir, "settings.json");

        /// <summary>
        /// Góc bẻ leader (độ). 0 = thẳng (mặc định).
        /// </summary>
        public static double AngleDegrees { get; set; }

        public static void Load()
        {
            try
            {
                if (!File.Exists(SettingsPath)) return;
                var json = File.ReadAllText(SettingsPath);
                var data = JsonSerializer.Deserialize<SettingsData>(json);
                if (data != null)
                    AngleDegrees = Math.Max(0, Math.Min(90, data.AngleDegrees));
            }
            catch { /* keep default */ }
        }

        public static void Save()
        {
            try
            {
                Directory.CreateDirectory(SettingsDir);
                var json = JsonSerializer.Serialize(new SettingsData { AngleDegrees = AngleDegrees });
                File.WriteAllText(SettingsPath, json);
            }
            catch { /* silently fail */ }
        }

        public static void Reset()
        {
            AngleDegrees = 0;
            Save();
        }

        private class SettingsData
        {
            public double AngleDegrees { get; set; }
        }
    }
}
