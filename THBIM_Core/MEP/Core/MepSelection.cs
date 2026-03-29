using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace THBIM.MEP.Core;

internal static class MepSelection
{
    /// <summary>Pick a single MEP element from the view.</summary>
    public static Element PickMepElement(UIDocument uidoc, string prompt)
    {
        var reference = uidoc.Selection.PickObject(ObjectType.Element,
            new ElementPickFilter(IsSupportedElement), prompt);
        return uidoc.Document.GetElement(reference)
               ?? throw new InvalidOperationException("No element selected.");
    }

    /// <summary>Pick a single element with a 3D pick point.</summary>
    public static (Element Element, XYZ PickPoint) PickMepElementWithPoint(UIDocument uidoc, string prompt)
    {
        var reference = uidoc.Selection.PickObject(ObjectType.PointOnElement,
            new ElementPickFilter(IsSupportedElement), prompt);
        var element = uidoc.Document.GetElement(reference)
                      ?? throw new InvalidOperationException("No element selected.");
        return (element, reference.GlobalPoint);
    }

    /// <summary>
    /// Get MEP elements from the current selection.
    /// If nothing selected, prompt user to select multiple.
    /// </summary>
    public static IList<Element> GetOrPickMepElements(UIDocument uidoc, string prompt)
    {
        var currentSelection = GetCurrentMepElements(uidoc);
        if (currentSelection.Count > 0)
            return currentSelection;

        var refs = uidoc.Selection.PickObjects(ObjectType.Element,
            new ElementPickFilter(IsSupportedElement), prompt);
        return refs.Select(r => uidoc.Document.GetElement(r))
            .Where(e => e is not null)
            .ToList()!;
    }

    /// <summary>Get MEP elements already in current selection.</summary>
    public static IList<Element> GetCurrentMepElements(UIDocument uidoc)
    {
        return uidoc.Selection.GetElementIds()
            .Select(id => uidoc.Document.GetElement(id))
            .Where(e => e is not null && IsSupportedElement(e))
            .ToList()!;
    }

    /// <summary>Get all elements from current selection (any type).</summary>
    public static IList<Element> GetCurrentElements(UIDocument uidoc)
    {
        return uidoc.Selection.GetElementIds()
            .Select(id => uidoc.Document.GetElement(id))
            .Where(x => x is not null)
            .Cast<Element>()
            .ToList();
    }

    public static bool IsSupportedElement(Element element)
        => element is MEPCurve or FamilyInstance or FabricationPart;

    public static bool IsFamilyInstance(Element element)
        => element is FamilyInstance fi && fi.MEPModel is not null;
}
