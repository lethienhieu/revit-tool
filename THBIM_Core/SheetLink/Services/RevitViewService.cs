using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace THBIM.Services
{
    public class RevitViewService
    {
        private readonly Document _doc;
        private readonly UIDocument _uiDoc;

        public RevitViewService(UIDocument uiDoc)
        {
            _uiDoc = uiDoc ?? throw new ArgumentNullException(nameof(uiDoc));
            _doc = uiDoc.Document;
        }

        public List<string> Get3DViewNames()
            => new FilteredElementCollector(_doc)
                .OfClass(typeof(View3D))
                .Cast<View3D>()
                .Where(v => !v.IsTemplate)
                .OrderBy(v => v.Name, StringComparer.OrdinalIgnoreCase)
                .Select(v => v.Name)
                .ToList();

        public string GetActive3DViewName()
            => _doc.ActiveView is View3D v && !v.IsTemplate ? v.Name : null;

        public void IsolateElements(IEnumerable<ElementId> elementIds)
        {
            var ids = elementIds?.Distinct().ToList();
            if (ids == null || !ids.Any())
                throw new ArgumentException("No elements selected for isolate.");

            var view = _doc.ActiveView ?? throw new InvalidOperationException("No active view.");
            if (!view.CanUseTemporaryVisibilityModes())
                throw new InvalidOperationException($"View '{view.Name}' does not support temporary isolate.");

            using var tx = new Transaction(_doc, "THBIM - Isolate");
            tx.Start();
            view.IsolateElementsTemporary(ids);
            tx.Commit();
        }

        public void ResetIsolate()
        {
            var view = _doc.ActiveView;
            if (view == null)
                return;

            using var tx = new Transaction(_doc, "THBIM - Reset Isolate");
            tx.Start();
            view.DisableTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate);
            tx.Commit();
        }

        public View3D CreateSectionBox(IEnumerable<ElementId> elementIds, double offsetMm = 500.0, bool duplicateActive = false)
        {
            var target = GetActive3DViewName();
            return CreateSectionBoxWithOptions(
                elementIds,
                target,
                offsetMm,
                duplicateActive,
                duplicateActive ? $"{target}_THBIM" : null,
                false);
        }

        public View3D CreateSectionBoxWithOptions(
            IEnumerable<ElementId> elementIds,
            string targetViewName,
            double offsetMm,
            bool duplicateView,
            string duplicateViewName,
            bool isolateElements)
        {
            var ids = elementIds?.Distinct().ToList();
            if (ids == null || !ids.Any())
                throw new ArgumentException("No elements selected.");

            var bbox = ComputeBoundingBox(ids, offsetMm);
            if (bbox == null)
                throw new InvalidOperationException("Cannot compute bounding box from selected elements.");

            using var tx = new Transaction(_doc, "THBIM - Section Box");
            tx.Start();

            var target = Resolve3DView(targetViewName);
            if (duplicateView)
            {
                var dupId = target.Duplicate(ViewDuplicateOption.Duplicate);
                target = _doc.GetElement(dupId) as View3D ?? target;
                if (!string.IsNullOrWhiteSpace(duplicateViewName))
                    target.Name = EnsureUniqueViewName(duplicateViewName.Trim());
            }

            target.SetSectionBox(bbox);
            target.IsSectionBoxActive = true;
            if (isolateElements && target.CanUseTemporaryVisibilityModes())
                target.IsolateElementsTemporary(ids);

            tx.Commit();
            _uiDoc.ActiveView = target;
            return target;
        }

        public List<ElementId> GetElementIdsByCategories(IEnumerable<string> categoryNames, IEnumerable<string> typeNames = null)
        {
            var ids = new List<ElementId>();
            var revit = new RevitDataService(_doc);
            var typeSet = typeNames != null
                ? new HashSet<string>(typeNames, StringComparer.OrdinalIgnoreCase)
                : null;

            foreach (var categoryName in categoryNames ?? Enumerable.Empty<string>())
            {
                var category = revit.GetCategoryByName(categoryName);
                if (category == null)
                    continue;

                var collector = new FilteredElementCollector(_doc)
                    .OfCategoryId(category.Id)
                    .WhereElementIsNotElementType();

                if (typeSet != null && typeSet.Any())
                {
                    ids.AddRange(collector.Where(el =>
                    {
                        var et = _doc.GetElement(el.GetTypeId()) as ElementType;
                        return et != null && typeSet.Contains($"{et.FamilyName} - {et.Name}");
                    }).Select(el => el.Id));
                }
                else
                {
                    ids.AddRange(collector.ToElementIds());
                }
            }

            return ids.Distinct().ToList();
        }

        private View3D Resolve3DView(string targetViewName)
        {
            if (!string.IsNullOrWhiteSpace(targetViewName))
            {
                var named = new FilteredElementCollector(_doc)
                    .OfClass(typeof(View3D))
                    .Cast<View3D>()
                    .FirstOrDefault(v => !v.IsTemplate &&
                                         string.Equals(v.Name, targetViewName, StringComparison.OrdinalIgnoreCase));
                if (named != null)
                    return named;
            }

            if (_doc.ActiveView is View3D active3D && !active3D.IsTemplate)
                return active3D;

            var default3D = new FilteredElementCollector(_doc)
                .OfClass(typeof(View3D))
                .Cast<View3D>()
                .FirstOrDefault(v => !v.IsTemplate &&
                                     string.Equals(v.Name, "{3D}", StringComparison.OrdinalIgnoreCase));
            if (default3D != null)
                return default3D;

            var viewType = new FilteredElementCollector(_doc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>()
                .FirstOrDefault(v => v.ViewFamily == ViewFamily.ThreeDimensional)
                ?? throw new InvalidOperationException("3D view type not found.");

            return View3D.CreateIsometric(_doc, viewType.Id);
        }

        private string EnsureUniqueViewName(string baseName)
        {
            var names = new HashSet<string>(
                new FilteredElementCollector(_doc)
                    .OfClass(typeof(View3D))
                    .Cast<View3D>()
                    .Select(v => v.Name),
                StringComparer.OrdinalIgnoreCase);

            if (!names.Contains(baseName))
                return baseName;

            var i = 1;
            while (names.Contains($"{baseName}_{i}"))
                i++;
            return $"{baseName}_{i}";
        }

        private BoundingBoxXYZ ComputeBoundingBox(List<ElementId> ids, double offsetMm)
        {
            var offsetFt = offsetMm / 304.8;
            var minX = double.MaxValue;
            var minY = double.MaxValue;
            var minZ = double.MaxValue;
            var maxX = double.MinValue;
            var maxY = double.MinValue;
            var maxZ = double.MinValue;
            var found = false;

            foreach (var id in ids)
            {
                var bb = _doc.GetElement(id)?.get_BoundingBox(null);
                if (bb == null)
                    continue;
                found = true;
                minX = Math.Min(minX, bb.Min.X);
                minY = Math.Min(minY, bb.Min.Y);
                minZ = Math.Min(minZ, bb.Min.Z);
                maxX = Math.Max(maxX, bb.Max.X);
                maxY = Math.Max(maxY, bb.Max.Y);
                maxZ = Math.Max(maxZ, bb.Max.Z);
            }

            if (!found)
                return null;

            return new BoundingBoxXYZ
            {
                Min = new XYZ(minX - offsetFt, minY - offsetFt, minZ - offsetFt),
                Max = new XYZ(maxX + offsetFt, maxY + offsetFt, maxZ + offsetFt)
            };
        }
    }
}
