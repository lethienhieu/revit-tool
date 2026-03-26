using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using THBIM.Models;

namespace THBIM.Services
{
    public class RevitDataService
    {
        private readonly Document _doc;
        private static readonly HashSet<string> DisciplineContainerNames =
            new(StringComparer.OrdinalIgnoreCase)
            {
                "<All Disciplines>",
                "Architecture",
                "Structure",
                "Mechanical",
                "Electrical",
                "Piping",
                "Infrastructure",
                "General",
                "MEP",
                "Analytical Model"
            };

        public RevitDataService(Document doc)
            => _doc = doc ?? throw new ArgumentNullException(nameof(doc));

        // ── Categories ────────────────────────────────────────────────────

        public List<CategoryItem> GetModelCategories()
        {
            if (_doc?.Settings?.Categories == null) return new();
            var result = new List<CategoryItem>();
            foreach (Category cat in _doc.Settings.Categories)
            {
                if (cat == null) continue;
                if (cat.CategoryType != CategoryType.Model)  continue;
                if (!cat.AllowsBoundParameters)              continue;
                if (DisciplineContainerNames.Contains(cat.Name)) continue;
                try
                {
                    if (!HasPlacedInstanceElements(cat)) continue;
                    result.Add(new CategoryItem(cat.Name, GetDiscipline(cat)));
                }
                catch
                {
                    // Some model categories cannot be collected by category id in certain documents.
                    continue;
                }
            }
            return result.OrderBy(c => c.Name).ToList();
        }

        public List<CategoryItem> GetAnnotationCategories()
        {
            if (_doc?.Settings?.Categories == null) return new();
            var result = new List<CategoryItem>();
            foreach (Category cat in _doc.Settings.Categories)
            {
                if (cat == null) continue;
                if (cat.CategoryType != CategoryType.Annotation) continue;
                if (!cat.AllowsBoundParameters)                  continue;
                try
                {
                    var count = new FilteredElementCollector(_doc)
                        .OfCategoryId(cat.Id)
                        .WhereElementIsNotElementType()
                        .GetElementCount();
                    if (count == 0) continue;
                    result.Add(new CategoryItem(cat.Name, "Annotation"));
                }
                catch
                {
                    continue;
                }
            }
            return result.OrderBy(c => c.Name).ToList();
        }

        // ── Parameters ────────────────────────────────────────────────────

        public List<ParameterItem> GetParameters(IEnumerable<string> categoryNames)
        {
            if (_doc == null || categoryNames == null) return new();
            var dict = new Dictionary<string, ParameterItem>(StringComparer.OrdinalIgnoreCase);
            foreach (var catName in categoryNames)
            {
                if (string.IsNullOrWhiteSpace(catName)) continue;
                if (DisciplineContainerNames.Contains(catName)) continue;

                var cat = GetCategoryByName(catName);
                if (cat == null) continue;

                try
                {
                    var inst = new FilteredElementCollector(_doc)
                        .OfCategoryId(cat.Id).WhereElementIsNotElementType().FirstElement();
                    if (inst != null) CollectParams(inst, dict, false);

                    var type = new FilteredElementCollector(_doc)
                        .OfCategoryId(cat.Id).WhereElementIsElementType().FirstElement();
                    if (type != null) CollectParams(type, dict, true);
                }
                catch
                {
                    continue;
                }
            }
            return dict.Values.OrderBy(p => p.Name).ToList();
        }

        private bool HasPlacedInstanceElements(Category cat)
        {
            if (cat == null) return false;

            try
            {
                foreach (var el in new FilteredElementCollector(_doc)
                    .OfCategoryId(cat.Id)
                    .WhereElementIsNotElementType()
                    .ToElements())
                {
                    if (el == null || el.ViewSpecific)
                        continue;
                    if (el.Category?.Id?.GetValue() != cat.Id.GetValue())
                        continue;

                    var hasSpatialPresence = el.Location != null;
                    if (!hasSpatialPresence)
                    {
                        try { hasSpatialPresence = el.get_BoundingBox(null) != null; }
                        catch { hasSpatialPresence = false; }
                    }

                    if (hasSpatialPresence)
                        return true;
                }
            }
            catch
            {
                // If this category cannot be queried in the current document, treat it as unsupported.
            }

            return false;
        }

        private void CollectParams(Element el,
                                    Dictionary<string, ParameterItem> dict,
                                    bool fromType)
        {
            foreach (Parameter p in el.Parameters)
            {
                if (p.Definition == null) continue;
                var name = p.Definition.Name;
                if (string.IsNullOrWhiteSpace(name) || dict.ContainsKey(name)) continue;
                var kind = p.IsReadOnly ? ParamKind.ReadOnly :
                           fromType    ? ParamKind.Type     :
                                         ParamKind.Instance;
                dict[name] = new ParameterItem(name, p.StorageType.ToString(), kind);
            }
        }

        // ── Elements ──────────────────────────────────────────────────────

        public List<CategoryItem> GetElementTypes(string categoryName)
        {
            if (_doc == null || string.IsNullOrWhiteSpace(categoryName)) return new();
            var cat = GetCategoryByName(categoryName);
            if (cat == null) return new();

            var placedTypeIds = new HashSet<long>();
            try
            {
                foreach (var el in new FilteredElementCollector(_doc)
                    .OfCategoryId(cat.Id)
                    .WhereElementIsNotElementType())
                {
                    if (el == null || el.ViewSpecific)
                        continue;
                    if (el.Category?.Id?.GetValue() != cat.Id.GetValue())
                        continue;

                    var tid = el.GetTypeId();
                    if (tid != null && tid != ElementId.InvalidElementId)
                        placedTypeIds.Add(tid.GetValue());
                }
            }
            catch
            {
                return new();
            }

            if (!placedTypeIds.Any())
                return new();

            return new FilteredElementCollector(_doc)
                .OfCategoryId(cat.Id).WhereElementIsElementType()
                .Cast<ElementType>()
                .Where(et => et.FamilyName != null && placedTypeIds.Contains(et.Id.GetValue()))
                .Select(et => new CategoryItem($"{et.FamilyName} - {et.Name}", categoryName))
                .OrderBy(c => c.Name).ToList();
        }

        public List<Element> GetElements(string categoryName,
                                          IEnumerable<string> typeNames = null)
        {
            if (_doc == null || string.IsNullOrWhiteSpace(categoryName)) return new();
            var cat = GetCategoryByName(categoryName);
            if (cat == null) return new();
            var elements = new FilteredElementCollector(_doc)
                .OfCategoryId(cat.Id).WhereElementIsNotElementType().ToElements().ToList();
            if (typeNames == null) return elements;
            var typeSet = new HashSet<string>(typeNames, StringComparer.OrdinalIgnoreCase);
            return elements.Where(el =>
            {
                var et = _doc.GetElement(el.GetTypeId()) as ElementType;
                return et != null && typeSet.Contains($"{et.FamilyName} - {et.Name}");
            }).ToList();
        }

        // ── Schedules ─────────────────────────────────────────────────────

        public List<ScheduleItem> GetSchedules()
        {
            if (_doc == null) return new();
            return new FilteredElementCollector(_doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .Where(v => !v.IsTemplate && !v.IsTitleblockRevisionSchedule)
                .Select(v => new ScheduleItem(v.Name, MapScheduleKind(v), v.Id.ToString()))
                .OrderBy(s => s.Name).ToList();
        }

        public List<ParameterItem> GetScheduleParameters(string scheduleElementId)
        {
            if (_doc == null || string.IsNullOrWhiteSpace(scheduleElementId))
                return new();

            if (!int.TryParse(scheduleElementId, out var idValue))
                return new();

            if (_doc.GetElement(new ElementId(idValue)) is not ViewSchedule schedule)
                return new();

            Element sample = null;
            try
            {
                sample = new FilteredElementCollector(_doc, schedule.Id)
                    .WhereElementIsNotElementType()
                    .FirstOrDefault();
            }
            catch
            {
                sample = null;
            }

            var result = new List<ParameterItem>();
            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var fieldId in schedule.Definition.GetFieldOrder())
            {
                ScheduleField field = null;
                try { field = schedule.Definition.GetField(fieldId); }
                catch { }
                if (field == null || field.IsHidden)
                    continue;

                var name = field?.GetName();
                if (string.IsNullOrWhiteSpace(name) || used.Contains(name))
                    continue;

                var kind = field.IsCalculatedField
                    ? ParamKind.ReadOnly
                    : ResolveScheduleFieldKind(sample, name);
                result.Add(new ParameterItem(name, "String", kind));
                used.Add(name);
            }

            return result;
        }

        private ParamKind ResolveScheduleFieldKind(Element sample, string fieldName)
        {
            if (sample == null || string.IsNullOrWhiteSpace(fieldName))
                return ParamKind.Instance;

            var instanceParam = sample.LookupParameter(fieldName);
            if (instanceParam != null)
                return instanceParam.IsReadOnly ? ParamKind.ReadOnly : ParamKind.Instance;

            var typeId = sample.GetTypeId();
            if (typeId != null && typeId != ElementId.InvalidElementId)
            {
                var typeParam = _doc.GetElement(typeId)?.LookupParameter(fieldName);
                if (typeParam != null)
                    return typeParam.IsReadOnly ? ParamKind.ReadOnly : ParamKind.Type;
            }

            return ParamKind.ReadOnly;
        }

        private static ScheduleKind MapScheduleKind(ViewSchedule vs)
        {
            if (vs.Definition.IsKeySchedule)     return ScheduleKind.KeySchedule;
            if (vs.Definition.IsMaterialTakeoff) return ScheduleKind.MaterialTakeoff;
            var catId = vs.Definition.CategoryId;
            if (catId == new ElementId(BuiltInCategory.OST_Sheets)) return ScheduleKind.SheetList;
            if (catId == new ElementId(BuiltInCategory.OST_Views))  return ScheduleKind.ViewList;
            return ScheduleKind.Regular;
        }

        // ── Spatial ───────────────────────────────────────────────────────

        public List<SpatialItem> GetRooms()
        {
            if (_doc == null) return new();
            return new FilteredElementCollector(_doc)
                .OfClass(typeof(SpatialElement))
                .OfCategory(BuiltInCategory.OST_Rooms)
                .Cast<Room>()
                .Where(r => r.Area > 0)
                .Select(r => new SpatialItem(
                    r.Id.GetValue(), r.Number, r.Name,
                    r.get_Parameter(BuiltInParameter.ROOM_PHASE)?.AsValueString() ?? "",
                    true))
                .OrderBy(r => r.Number).ToList();
        }

        public List<SpatialItem> GetSpaces()
        {
            if (_doc == null) return new();
            return new FilteredElementCollector(_doc)
                .OfClass(typeof(SpatialElement))
                .OfCategory(BuiltInCategory.OST_MEPSpaces)
                .Cast<Space>()
                .Where(s => s.Area > 0)
                .Select(s => new SpatialItem(
                    s.Id.GetValue(), s.Number, s.Name,
                    s.get_Parameter(BuiltInParameter.ROOM_PHASE)?.AsValueString() ?? "",
                    false))
                .OrderBy(s => s.Number).ToList();
        }

        // ── Helpers ───────────────────────────────────────────────────────

        public List<ParameterItem> GetSpatialParameters(bool rooms)
        {
            if (_doc == null) return new();

            var dict = new Dictionary<string, ParameterItem>(StringComparer.OrdinalIgnoreCase);
            var bic = rooms ? BuiltInCategory.OST_Rooms : BuiltInCategory.OST_MEPSpaces;

            try
            {
                var instance = new FilteredElementCollector(_doc)
                    .OfCategory(bic)
                    .WhereElementIsNotElementType()
                    .FirstElement();
                if (instance != null)
                    CollectParams(instance, dict, false);

                var type = instance != null ? _doc.GetElement(instance.GetTypeId()) : null;
                if (type != null)
                    CollectParams(type, dict, true);
            }
            catch
            {
                return new();
            }

            return dict.Values.OrderBy(p => p.Name).ToList();
        }

        public Category GetCategoryByName(string name)
        {
            if (_doc?.Settings?.Categories == null || string.IsNullOrWhiteSpace(name))
                return null;
            foreach (Category cat in _doc.Settings.Categories)
                if (string.Equals(cat.Name, name, StringComparison.OrdinalIgnoreCase))
                    return cat;
            return null;
        }

        private static string GetDiscipline(Category cat)
        {
            if (cat == null) return "General";

            try
            {
                var bic = (BuiltInCategory)(int)cat.Id.GetValue();
                return bic switch
                {
                    BuiltInCategory.OST_StructuralColumns or
                    BuiltInCategory.OST_StructuralFraming or
                    BuiltInCategory.OST_StructuralFoundation or
                    BuiltInCategory.OST_StructConnections or
                    BuiltInCategory.OST_Rebar or
                    BuiltInCategory.OST_AreaRein or
                    BuiltInCategory.OST_PathRein or
                    BuiltInCategory.OST_AnalyticalNodes => "Structure",

                    BuiltInCategory.OST_DuctCurves or
                    BuiltInCategory.OST_DuctFitting or
                    BuiltInCategory.OST_DuctAccessory or
                    BuiltInCategory.OST_MechanicalEquipment or
                    BuiltInCategory.OST_MEPSpaces or
                    BuiltInCategory.OST_HVAC_Zones => "Mechanical",

                    BuiltInCategory.OST_ElectricalEquipment or
                    BuiltInCategory.OST_ElectricalFixtures or
                    BuiltInCategory.OST_CableTray or
                    BuiltInCategory.OST_CableTrayFitting or
                    BuiltInCategory.OST_Conduit or
                    BuiltInCategory.OST_ConduitFitting or
                    BuiltInCategory.OST_LightingDevices or
                    BuiltInCategory.OST_LightingFixtures => "Electrical",

                    BuiltInCategory.OST_PipeCurves or
                    BuiltInCategory.OST_PipeFitting or
                    BuiltInCategory.OST_PipeAccessory or
                    BuiltInCategory.OST_PlumbingFixtures or
                    BuiltInCategory.OST_Sprinklers => "Piping",

                    BuiltInCategory.OST_Topography or
                    BuiltInCategory.OST_Site => "Infrastructure",

                    BuiltInCategory.OST_Walls or
                    BuiltInCategory.OST_Doors or
                    BuiltInCategory.OST_Windows or
                    BuiltInCategory.OST_Roofs or
                    BuiltInCategory.OST_Floors or
                    BuiltInCategory.OST_Ceilings or
                    BuiltInCategory.OST_Stairs or
                    BuiltInCategory.OST_Ramps or
                    BuiltInCategory.OST_Railings or
                    BuiltInCategory.OST_Rooms or
                    BuiltInCategory.OST_Furniture or
                    BuiltInCategory.OST_CurtainWallPanels => "Architecture",

                    _ => InferDisciplineFromName(cat.Name)
                };
            }
            catch
            {
                return InferDisciplineFromName(cat.Name);
            }
        }

        private static string InferDisciplineFromName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "General";
            var n = name.ToLowerInvariant();

            if (n.Contains("analytical") || n.Contains("struct") || n.Contains("rebar"))
                return "Structure";
            if (n.Contains("mechanical") || n.Contains("hvac") || n.Contains("duct"))
                return "Mechanical";
            if (n.Contains("electrical") || n.Contains("cable") || n.Contains("conduit") || n.Contains("lighting"))
                return "Electrical";
            if (n.Contains("pipe") || n.Contains("plumb") || n.Contains("sprinkler"))
                return "Piping";
            if (n.Contains("topo") || n.Contains("site") || n.Contains("road") || n.Contains("bridge"))
                return "Infrastructure";
            if (n.Contains("wall") || n.Contains("door") || n.Contains("window") || n.Contains("room") || n.Contains("floor") || n.Contains("roof"))
                return "Architecture";

            return "General";
        }
    }
}
