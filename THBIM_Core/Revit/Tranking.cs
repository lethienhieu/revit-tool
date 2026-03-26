using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
#nullable disable


namespace THBIM
{
    // ========= 1) Extensible Storage: KHÔNG yêu cầu Unit =========
    internal static class SleeveTracking
    {
        // QUAN TRỌNG: Hãy đổi GUID này sang một GUID mới hoàn toàn để tránh xung đột với bản test bị lỗi trước đó
        private static readonly Guid _schemaGuid = new Guid("AAAAAAAA-BBBB-CCCC-DDDD-EEEEFFFF0001"); // <--- GEN GUID MỚI VÀ PASTE VÀO ĐÂY

        private const string SchemaName = "TH_SleeveTracking_v2025"; // Đổi tên nhẹ để chắc chắn
        private const string Field_BaselineS = "BaselineStr";
        private const string Field_HostId = "HostId";
        private const string Field_NeedsRev = "NeedsReview";
        private const string Field_ChangedAt = "ChangedAtUtc";

        public static Schema GetOrCreateSchema()
        {
            var s = Schema.Lookup(_schemaGuid);
            if (s != null) return s;

            var sb = new SchemaBuilder(_schemaGuid);
            sb.SetSchemaName(SchemaName);

            sb.AddSimpleField(Field_BaselineS, typeof(string));

            // [FIX 1] Đổi typeof(int) thành typeof(long) cho Revit 2025
            sb.AddSimpleField(Field_HostId, typeof(long));

            sb.AddSimpleField(Field_NeedsRev, typeof(bool));
            sb.AddSimpleField(Field_ChangedAt, typeof(string));

            return sb.Finish();
        }

        // ---- Helpers: XYZ <-> string (InvariantCulture) ----
        private static string XYZToStr(XYZ p) =>
            string.Format(CultureInfo.InvariantCulture, "{0:R}|{1:R}|{2:R}", p.X, p.Y, p.Z);

        private static XYZ StrToXYZ(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return null;
            var parts = s.Split('|');
            if (parts.Length != 3) return null;
            if (double.TryParse(parts[0], NumberStyles.Float, CultureInfo.InvariantCulture, out double x) &&
                double.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out double y) &&
                double.TryParse(parts[2], NumberStyles.Float, CultureInfo.InvariantCulture, out double z))
            {
                return new XYZ(x, y, z);
            }
            return null;
        }

        public static void SnapshotBaseline(Document doc, FamilyInstance inst)
        {
            var s = GetOrCreateSchema();
            var ent = new Entity(s);

            ent.Set(s.GetField(Field_BaselineS), XYZToStr(GetInstanceOrigin(inst)));
            ent.Set(s.GetField(Field_HostId), inst.Host?.Id.GetValue() ?? -1);
            ent.Set(s.GetField(Field_NeedsRev), false);
            ent.Set(s.GetField(Field_ChangedAt), DateTime.UtcNow.ToString("o"));

            inst.SetEntity(ent);
        }

        public static void MarkChanged(FamilyInstance inst)
        {
            var s = GetOrCreateSchema();
            var ent = inst.GetEntity(s);
            if (!ent.IsValid()) ent = new Entity(s);

            ent.Set(s.GetField(Field_NeedsRev), true);
            ent.Set(s.GetField(Field_ChangedAt), DateTime.UtcNow.ToString("o"));

            inst.SetEntity(ent);
        }

        public static void AcceptChange(FamilyInstance inst)
        {
            var s = GetOrCreateSchema();
            var ent = inst.GetEntity(s);
            if (!ent.IsValid()) ent = new Entity(s);

            ent.Set(s.GetField(Field_BaselineS), XYZToStr(GetInstanceOrigin(inst)));
            ent.Set(s.GetField(Field_NeedsRev), false);
            ent.Set(s.GetField(Field_ChangedAt), DateTime.UtcNow.ToString("o"));

            inst.SetEntity(ent);
        }

        public static bool NeedsReview(FamilyInstance inst)
        {
            var s = GetOrCreateSchema();
            var e = inst.GetEntity(s);
            return e.IsValid() && e.Get<bool>(s.GetField(Field_NeedsRev));
        }

        public static XYZ GetBaseline(FamilyInstance inst)
        {
            var s = GetOrCreateSchema();
            var e = inst.GetEntity(s);
            if (!e.IsValid()) return null;
            return StrToXYZ(e.Get<string>(s.GetField(Field_BaselineS)));
        }

        public static XYZ GetInstanceOrigin(FamilyInstance inst)
        {
            var tr = inst.GetTransform();
            if (tr != null) return tr.Origin;
            if (inst.Location is LocationPoint lp) return lp.Point;
            var bb = inst.get_BoundingBox(null);
            return bb != null ? 0.5 * (bb.Min + bb.Max) : XYZ.Zero;
        }
    }

    // ========= 2) Updater: đánh dấu khi host/sleeve thay đổi =========
    public class SleeveChangeUpdater : IUpdater
    {
        private static readonly UpdaterId _uid =
            new UpdaterId(new AddInId(new Guid("8B8E1B86-E9F7-42C1-B7A6-7A3B8A0A7F3F")),   // trùng AddInId trong .addin
                          new Guid("A1D8F0A4-6D3D-44A4-8C93-3F9F1E8F1A20"));

        public UpdaterId GetUpdaterId() => _uid;
        public string GetUpdaterName() => "TH Sleeve Change Tracker";
        public string GetAdditionalInformation() => "Marks hosted sleeves as NeedsReview when host MEP/sleeve changes.";
        public ChangePriority GetChangePriority() => ChangePriority.MEPSystems;

        public void Execute(UpdaterData data)
        {
            var doc = data.GetDocument();

            // a) Chính sleeve thay đổi
            foreach (var id in data.GetModifiedElementIds())
            {
                var e = doc.GetElement(id);
                if (IsSleeve(e, out var inst)) SleeveTracking.MarkChanged(inst);
            }

            // b) Host MEP thay đổi
            var modifiedHosts = data.GetModifiedElementIds()
                .Select(id => doc.GetElement(id))
                .Where(e => e is Pipe || e is Duct || e is CableTray || e is Conduit)
                .Select(e => e.Id)
                .ToHashSet();

            if (modifiedHosts.Count == 0) return;

            var sleeves = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(IsSleeveName)
                .ToList();

            foreach (var s in sleeves)
            {
                var host = s.Host as Element;
                if (host != null && modifiedHosts.Contains(host.Id))
                    SleeveTracking.MarkChanged(s);
            }
        }

        private static bool IsSleeve(Element e, out FamilyInstance fi)
        {
            fi = e as FamilyInstance;
            return fi != null && IsSleeveName(fi);
        }

        public static bool IsSleeveName(FamilyInstance fi)
        {
            var fam = fi.Symbol?.Family;
            if (fam == null) return false;
            string n = fam.Name ?? "";
            return n.Equals("TH_ROUND_SLEEVE", StringComparison.OrdinalIgnoreCase)
                || n.Equals("TH_RECTANGULAR_SLEEVE", StringComparison.OrdinalIgnoreCase)
                || n.Equals("TH_RECTANGULAR_MULTI_SLEEVE", StringComparison.OrdinalIgnoreCase);
        }

        // Gọi từ App.OnStartup
        public static void Register(UIControlledApplication app)
        {
            try { UpdaterRegistry.UnregisterUpdater(_uid); } catch { }

            var up = new SleeveChangeUpdater();
            UpdaterRegistry.RegisterUpdater(up);

            // Host MEP
            var hostFilter = new LogicalOrFilter(new List<ElementFilter>
            {
                new ElementClassFilter(typeof(Pipe)),
                new ElementClassFilter(typeof(Duct)),
                new ElementClassFilter(typeof(CableTray)),
                new ElementClassFilter(typeof(Conduit))
            });
            UpdaterRegistry.AddTrigger(_uid, hostFilter, Element.GetChangeTypeAny());

            // Bản thân sleeve (FI)
            var fiFilter = new ElementClassFilter(typeof(FamilyInstance));
            UpdaterRegistry.AddTrigger(_uid, fiFilter, Element.GetChangeTypeGeometry());
            // ví dụ thêm trigger param (Comments)
            UpdaterRegistry.AddTrigger(
                _uid, fiFilter,
                Element.GetChangeTypeParameter(new ElementId(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)));
        }
        public static void Unregister()
        {
            try { UpdaterRegistry.UnregisterUpdater(_uid); } catch { /* ignore */ }
        }
    }

    // ========= 3) Commands: chọn / accept =========
    [Transaction(TransactionMode.Manual)]
    public class SelectChanged : IExternalCommand
    {
        public Result Execute(ExternalCommandData cd, ref string msg, ElementSet set)
        {
            if (!THBIM.Licensing.LicenseManager.EnsureActivated(null))
            {
                return Result.Cancelled;
            }
            if (!THBIM.Licensing.LicenseManager.EnsurePremium())
                return Result.Cancelled;

            var uidoc = cd.Application.ActiveUIDocument;
            var doc = uidoc.Document;

            var ids = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(SleeveChangeUpdater.IsSleeveName)
                .Where(SleeveTracking.NeedsReview)
                .Select(x => x.Id)
                .ToList();

            uidoc.Selection.SetElementIds(ids);
            TaskDialog.Show("OPENING", $"Found {ids.Count} opening(s) changed. (Selected)");
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class Accept : IExternalCommand
    {
        public Result Execute(ExternalCommandData cd, ref string msg, ElementSet set)
        {

            if (!THBIM.Licensing.LicenseManager.EnsureActivated(null))
            {
                return Result.Cancelled;
            }
            if (!THBIM.Licensing.LicenseManager.EnsurePremium())
                return Result.Cancelled;

            var doc = cd.Application.ActiveUIDocument.Document;

            var sleeves = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(SleeveChangeUpdater.IsSleeveName)
                .Where(SleeveTracking.NeedsReview)
                .ToList();

            using (var t = new Transaction(doc, "Accept Sleeve Changes"))
            {
                t.Start();
                foreach (var s in sleeves) SleeveTracking.AcceptChange(s);
                t.Commit();
            }

            TaskDialog.Show("OPENING", $"Accepted {sleeves.Count} opening(s). Baseline updated.");
            return Result.Succeeded;
        }
    }
}
