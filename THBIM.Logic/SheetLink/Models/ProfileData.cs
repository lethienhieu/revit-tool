using System.Collections.Generic;

namespace THBIM.Models
{
    /// <summary>
    /// Profile data — stores the full configuration of a SheetLink working session.
    /// </summary>
    public class ProfileData
    {
        // ── Metadata ─────────────────────────────────────────────────────

        public string Name { get; set; } = "New Profile";
        public string CreatedAt { get; set; }
        public string UpdatedAt { get; set; }

        // ── Tab: Model Categories ─────────────────────────────────────────

        public List<string> ModelCategories { get; set; } = new();
        public List<string> ModelParameters { get; set; } = new();

        /// <summary>Scope: "Whole" / "Active" / "Current"</summary>
        public string ModelScope { get; set; } = "Whole";

        public bool ModelIncludeLinkedFiles { get; set; } = false;
        public bool ModelExportByTypeId { get; set; } = false;

        // ── Tab: Annotation Categories ────────────────────────────────────

        public List<string> AnnotationCategories { get; set; } = new();
        public List<string> AnnotationParameters { get; set; } = new();
        public string AnnotationScope { get; set; } = "Whole";
        public bool AnnotationIncludeLinked { get; set; } = false;
        public bool AnnotationExportByTypeId { get; set; } = false;

        // ── Tab: Elements ─────────────────────────────────────────────────

        public string ElementsCategory { get; set; }
        public List<string> ElementsSelected { get; set; } = new();
        public List<string> ElementsParameters { get; set; } = new();
        public string ElementsScope { get; set; } = "Whole";
        public bool ElementsIncludeLinked { get; set; } = false;
        public bool ElementsExportByTypeId { get; set; } = false;

        // ── Tab: Schedules ────────────────────────────────────────────────

        public List<string> Schedules { get; set; } = new();
        public List<string> ScheduleParameters { get; set; } = new();
        public bool ScheduleExportByTypeId { get; set; } = false;
        public string ScheduleScope { get; set; } = "Whole";

        // ── Tab: Spatial ──────────────────────────────────────────────────

        public string SpatialType { get; set; } = "Rooms";
        public List<long> SpatialSelected { get; set; } = new();
        public List<string> SpatialParameters { get; set; } = new();
        public bool SpatialIncludeLinked { get; set; } = false;

        // ── Export settings ───────────────────────────────────────────────

        public string LastExcelPath { get; set; }
        public string LastGoogleDriveFolderId { get; set; }
    }
}
