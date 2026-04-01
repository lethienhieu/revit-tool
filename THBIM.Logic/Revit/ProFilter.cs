using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace THBIM
{
    // ===== FILTER BY CATEGORY (giữ nguyên) =====
    [Transaction(TransactionMode.Manual)]
    public class ProFilterCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            if (!Licensing.LicenseManager.EnsureActivated(null)) return Result.Cancelled;
            if (!Licensing.LicenseManager.EnsurePremium()) return Result.Cancelled;

            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                Reference pickedRef;
                try { pickedRef = uidoc.Selection.PickObject(ObjectType.Element, "Pick a source element to filter by Category"); }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException) { return Result.Cancelled; }

                Element source = doc.GetElement(pickedRef);
                if (source?.Category == null)
                {
                    TaskDialog.Show("ProFilter", "Selected element has no valid Category.");
                    return Result.Failed;
                }

                string label = source.Category.Name;
                var filter = new ProCategoryFilter(source.Category.Id);
                return DoRectangleSelect(uidoc, filter, label);
            }
            catch (Exception ex) { message = ex.Message; return Result.Failed; }
        }

        internal static Result DoRectangleSelect(UIDocument uidoc, ISelectionFilter filter, string label)
        {
            try
            {
                var picked = uidoc.Selection.PickElementsByRectangle(filter, $"Window-select to pick all [{label}]");
                var ids = picked.Select(e => e.Id).ToList();

                if (ids.Count > 0)
                    uidoc.Selection.SetElementIds(ids);
                else
                    TaskDialog.Show("ProFilter", $"No [{label}] found in selection region.");
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            return Result.Succeeded;
        }
    }

    // ===== FILTER BY FAMILY NAME =====
    [Transaction(TransactionMode.Manual)]
    public class ProFilterByFamilyCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            if (!Licensing.LicenseManager.EnsureActivated(null)) return Result.Cancelled;
            if (!Licensing.LicenseManager.EnsurePremium()) return Result.Cancelled;

            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                Reference pickedRef;
                try { pickedRef = uidoc.Selection.PickObject(ObjectType.Element, "Pick a source element to filter by Family"); }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException) { return Result.Cancelled; }

                Element source = doc.GetElement(pickedRef);
                string familyName = GetFamilyName(source);

                if (string.IsNullOrEmpty(familyName))
                {
                    TaskDialog.Show("ProFilter", "Cannot determine Family name from selected element.");
                    return Result.Failed;
                }

                var filter = new ProFamilyFilter(familyName);
                return ProFilterCommand.DoRectangleSelect(uidoc, filter, familyName);
            }
            catch (Exception ex) { message = ex.Message; return Result.Failed; }
        }

        private static string GetFamilyName(Element elem)
        {
            if (elem is FamilyInstance fi) return fi.Symbol?.Family?.Name;
            // System families: use Category name as fallback
            return elem?.Category?.Name;
        }
    }

    // ===== FILTER BY TYPE NAME =====
    [Transaction(TransactionMode.Manual)]
    public class ProFilterByTypeCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            if (!Licensing.LicenseManager.EnsureActivated(null)) return Result.Cancelled;
            if (!Licensing.LicenseManager.EnsurePremium()) return Result.Cancelled;

            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                Reference pickedRef;
                try { pickedRef = uidoc.Selection.PickObject(ObjectType.Element, "Pick a source element to filter by Type"); }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException) { return Result.Cancelled; }

                Element source = doc.GetElement(pickedRef);
                if (source?.Category == null)
                {
                    TaskDialog.Show("ProFilter", "Selected element has no valid Category.");
                    return Result.Failed;
                }

                string typeName = source.Name;
                ElementId catId = source.Category.Id;

                if (string.IsNullOrEmpty(typeName))
                {
                    TaskDialog.Show("ProFilter", "Cannot determine Type name from selected element.");
                    return Result.Failed;
                }

                var filter = new ProTypeFilter(catId, typeName);
                return ProFilterCommand.DoRectangleSelect(uidoc, filter, typeName);
            }
            catch (Exception ex) { message = ex.Message; return Result.Failed; }
        }
    }

    // ===== SELECTION FILTERS =====

    public class ProCategoryFilter : ISelectionFilter
    {
        private readonly ElementId _targetCategoryId;
        public ProCategoryFilter(ElementId targetCategoryId) { _targetCategoryId = targetCategoryId; }
        public bool AllowElement(Element elem) => elem.Category != null && elem.Category.Id == _targetCategoryId;
        public bool AllowReference(Reference reference, XYZ position) => false;
    }

    public class ProFamilyFilter : ISelectionFilter
    {
        private readonly string _targetFamilyName;
        public ProFamilyFilter(string targetFamilyName) { _targetFamilyName = targetFamilyName; }

        public bool AllowElement(Element elem)
        {
            if (elem is FamilyInstance fi)
                return fi.Symbol?.Family?.Name == _targetFamilyName;
            // System families: match by Category name
            return elem.Category?.Name == _targetFamilyName;
        }

        public bool AllowReference(Reference reference, XYZ position) => false;
    }

    public class ProTypeFilter : ISelectionFilter
    {
        private readonly ElementId _targetCategoryId;
        private readonly string _targetTypeName;

        public ProTypeFilter(ElementId targetCategoryId, string targetTypeName)
        {
            _targetCategoryId = targetCategoryId;
            _targetTypeName = targetTypeName;
        }

        public bool AllowElement(Element elem)
        {
            return elem.Category != null
                && elem.Category.Id == _targetCategoryId
                && elem.Name == _targetTypeName;
        }

        public bool AllowReference(Reference reference, XYZ position) => false;
    }
}
