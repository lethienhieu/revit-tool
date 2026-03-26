using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using THBIM.Models;

namespace THBIM.Services
{
    public class ProfileManager
    {
        private static ProfileManager _instance;
        public  static ProfileManager Instance => _instance ??= new ProfileManager();
        private ProfileManager() { }

        public string ProfileFolder { get; set; } =
            Path.Combine(Environment.GetFolderPath(
                Environment.SpecialFolder.ApplicationData),
                "THBIM", "SheetLink", "Profiles");

        // ── Save / Load (TXT format) ─────────────────────────────────────

        public string Save(ProfileData profile)
        {
            if (string.IsNullOrWhiteSpace(profile.Name))
                throw new ArgumentException("Profile name cannot be empty.");
            EnsureFolder();
            var now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            if (string.IsNullOrEmpty(profile.CreatedAt)) profile.CreatedAt = now;
            profile.UpdatedAt = now;
            var path = GetFilePath(profile.Name);
            File.WriteAllText(path, SerializeToTxt(profile));
            return path;
        }

        public ProfileData Load(string name)
        {
            var path = GetFilePath(name);
            if (!File.Exists(path)) return null;
            return DeserializeFromTxt(File.ReadAllText(path));
        }

        public ProfileData LoadFromFile(string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"File not found: {filePath}");
            return DeserializeFromTxt(File.ReadAllText(filePath));
        }

        public void ExportToFile(ProfileData profile, string targetPath)
            => File.WriteAllText(targetPath, SerializeToTxt(profile));

        public List<string> GetAllNames()
        {
            EnsureFolder();
            return Directory.GetFiles(ProfileFolder, "*.txt")
                .Select(f =>
                {
                    try
                    {
                        var p = DeserializeFromTxt(File.ReadAllText(f));
                        return (Name: p?.Name ?? Path.GetFileNameWithoutExtension(f),
                                Updated: p?.UpdatedAt ?? string.Empty);
                    }
                    catch { return (Name: Path.GetFileNameWithoutExtension(f), Updated: string.Empty); }
                })
                .OrderByDescending(x => x.Updated)
                .Select(x => x.Name)
                .ToList();
        }

        public void Delete(string name)
        {
            var path = GetFilePath(name);
            if (File.Exists(path)) File.Delete(path);
        }

        public bool Exists(string name) => File.Exists(GetFilePath(name));

        public void ApplyToState(ProfileData profile)
        {
            SheetLinkState.Instance.ActiveProfile =
                profile ?? throw new ArgumentNullException(nameof(profile));
        }

        public ProfileData SnapshotFromState(string name)
        {
            var s       = SheetLinkState.Instance;
            var profile = s.ActiveProfile ?? new ProfileData();
            profile.Name = name;

            profile.ModelCategories  = s.ModelCategories.Where(c => c.IsChecked).Select(c => c.Name).ToList();
            profile.ModelParameters  = s.ModelSelectedParams.Select(p => p.Name).ToList();
            profile.AnnotationCategories = s.AnnCategories.Where(c => c.IsChecked).Select(c => c.Name).ToList();
            profile.AnnotationParameters = s.AnnSelectedParams.Select(p => p.Name).ToList();
            profile.ElementsCategory = s.ElemCategories.FirstOrDefault(c => c.IsChecked)?.Name;
            profile.ElementsParameters = s.ElemSelectedParams.Select(p => p.Name).ToList();
            profile.Schedules        = s.Schedules.Where(sc => sc.IsChecked).Select(sc => sc.Name).ToList();
            profile.ScheduleParameters = s.SchedParams.Select(p => p.Name).ToList();
            profile.SpatialParameters = s.SpatialSelParams.Select(p => p.Name).ToList();
            profile.SpatialSelected  = s.SpatialItems.Where(i => i.IsChecked).Select(i => i.ElementId).ToList();

            return profile;
        }

        // ── TXT Serialization (public for RevitStorageService) ────────────

        public string SerializeProfile(ProfileData p) => SerializeToTxt(p);
        public ProfileData DeserializeProfile(string text) => DeserializeFromTxt(text);

        private static string SerializeToTxt(ProfileData p)
        {
            var lines = new List<string>
            {
                $"Name={Esc(p.Name)}",
                $"CreatedAt={Esc(p.CreatedAt)}",
                $"UpdatedAt={Esc(p.UpdatedAt)}",
                $"ModelCategories={JoinList(p.ModelCategories)}",
                $"ModelParameters={JoinList(p.ModelParameters)}",
                $"ModelScope={Esc(p.ModelScope)}",
                $"ModelIncludeLinkedFiles={p.ModelIncludeLinkedFiles}",
                $"ModelExportByTypeId={p.ModelExportByTypeId}",
                $"AnnotationCategories={JoinList(p.AnnotationCategories)}",
                $"AnnotationParameters={JoinList(p.AnnotationParameters)}",
                $"AnnotationScope={Esc(p.AnnotationScope)}",
                $"AnnotationIncludeLinked={p.AnnotationIncludeLinked}",
                $"AnnotationExportByTypeId={p.AnnotationExportByTypeId}",
                $"ElementsCategory={Esc(p.ElementsCategory)}",
                $"ElementsSelected={JoinList(p.ElementsSelected)}",
                $"ElementsParameters={JoinList(p.ElementsParameters)}",
                $"ElementsScope={Esc(p.ElementsScope)}",
                $"ElementsIncludeLinked={p.ElementsIncludeLinked}",
                $"ElementsExportByTypeId={p.ElementsExportByTypeId}",
                $"Schedules={JoinList(p.Schedules)}",
                $"ScheduleParameters={JoinList(p.ScheduleParameters)}",
                $"ScheduleExportByTypeId={p.ScheduleExportByTypeId}",
                $"ScheduleScope={Esc(p.ScheduleScope)}",
                $"SpatialType={Esc(p.SpatialType)}",
                $"SpatialSelected={JoinIntList(p.SpatialSelected)}",
                $"SpatialParameters={JoinList(p.SpatialParameters)}",
                $"SpatialIncludeLinked={p.SpatialIncludeLinked}",
                $"LastExcelPath={Esc(p.LastExcelPath)}",
                $"LastGoogleDriveFolderId={Esc(p.LastGoogleDriveFolderId)}"
            };
            return string.Join(Environment.NewLine, lines);
        }

        private static ProfileData DeserializeFromTxt(string text)
        {
            var p = new ProfileData();
            if (string.IsNullOrWhiteSpace(text)) return p;

            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var line in text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var idx = line.IndexOf('=');
                if (idx <= 0) continue;
                dict[line[..idx].Trim()] = line[(idx + 1)..];
            }

            p.Name = Get(dict, "Name", "New Profile");
            p.CreatedAt = Get(dict, "CreatedAt");
            p.UpdatedAt = Get(dict, "UpdatedAt");
            p.ModelCategories = GetList(dict, "ModelCategories");
            p.ModelParameters = GetList(dict, "ModelParameters");
            p.ModelScope = Get(dict, "ModelScope", "Whole");
            p.ModelIncludeLinkedFiles = GetBool(dict, "ModelIncludeLinkedFiles");
            p.ModelExportByTypeId = GetBool(dict, "ModelExportByTypeId");
            p.AnnotationCategories = GetList(dict, "AnnotationCategories");
            p.AnnotationParameters = GetList(dict, "AnnotationParameters");
            p.AnnotationScope = Get(dict, "AnnotationScope", "Whole");
            p.AnnotationIncludeLinked = GetBool(dict, "AnnotationIncludeLinked");
            p.AnnotationExportByTypeId = GetBool(dict, "AnnotationExportByTypeId");
            p.ElementsCategory = Get(dict, "ElementsCategory");
            p.ElementsSelected = GetList(dict, "ElementsSelected");
            p.ElementsParameters = GetList(dict, "ElementsParameters");
            p.ElementsScope = Get(dict, "ElementsScope", "Whole");
            p.ElementsIncludeLinked = GetBool(dict, "ElementsIncludeLinked");
            p.ElementsExportByTypeId = GetBool(dict, "ElementsExportByTypeId");
            p.Schedules = GetList(dict, "Schedules");
            p.ScheduleParameters = GetList(dict, "ScheduleParameters");
            p.ScheduleExportByTypeId = GetBool(dict, "ScheduleExportByTypeId");
            p.ScheduleScope = Get(dict, "ScheduleScope", "Whole");
            p.SpatialType = Get(dict, "SpatialType", "Rooms");
            p.SpatialSelected = GetIntList(dict, "SpatialSelected");
            p.SpatialParameters = GetList(dict, "SpatialParameters");
            p.SpatialIncludeLinked = GetBool(dict, "SpatialIncludeLinked");
            p.LastExcelPath = Get(dict, "LastExcelPath");
            p.LastGoogleDriveFolderId = Get(dict, "LastGoogleDriveFolderId");

            return p;
        }

        private static string Esc(string v) => v ?? string.Empty;
        private static string JoinList(List<string> list) =>
            list != null ? string.Join("|", list.Where(s => !string.IsNullOrEmpty(s))) : string.Empty;
        private static string JoinIntList(List<long> list) =>
            list != null ? string.Join("|", list) : string.Empty;

        private static string Get(Dictionary<string, string> d, string key, string def = null) =>
            d.TryGetValue(key, out var v) && !string.IsNullOrEmpty(v) ? v : def ?? string.Empty;

        private static bool GetBool(Dictionary<string, string> d, string key) =>
            d.TryGetValue(key, out var v) && bool.TryParse(v, out var b) && b;

        private static List<string> GetList(Dictionary<string, string> d, string key)
        {
            if (!d.TryGetValue(key, out var v) || string.IsNullOrWhiteSpace(v))
                return new List<string>();
            return v.Split('|').Where(s => !string.IsNullOrEmpty(s)).ToList();
        }

        private static List<long> GetIntList(Dictionary<string, string> d, string key)
        {
            if (!d.TryGetValue(key, out var v) || string.IsNullOrWhiteSpace(v))
                return new List<long>();
            return v.Split('|')
                .Where(s => !string.IsNullOrEmpty(s))
                .Select(s => long.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var i) ? i : -1L)
                .Where(i => i >= 0)
                .ToList();
        }

        // ── File path helpers ────────────────────────────────────────────

        private string GetFilePath(string name)
        {
            var safe = new string(name.Select(c =>
                Path.GetInvalidFileNameChars().Contains(c) ? '_' : c).ToArray());
            return Path.Combine(ProfileFolder, safe + ".txt");
        }

        private void EnsureFolder()
        {
            if (!Directory.Exists(ProfileFolder))
                Directory.CreateDirectory(ProfileFolder);
        }
    }
}
