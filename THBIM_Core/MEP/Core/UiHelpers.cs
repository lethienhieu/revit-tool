using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace THBIM.MEP.Core;

internal static class UiHelpers
{
    public static void ShowInfo(string title, string message)
    {
        var dialog = new TaskDialog(title)
        {
            MainInstruction = title,
            MainContent = message,
            CommonButtons = TaskDialogCommonButtons.Ok
        };
        dialog.Show();
    }

    public static void ShowWarning(string title, string message)
    {
        var dialog = new TaskDialog(title)
        {
            MainInstruction = title,
            MainContent = message,
            MainIcon = TaskDialogIcon.TaskDialogIconWarning,
            CommonButtons = TaskDialogCommonButtons.Ok
        };
        dialog.Show();
    }

    public static XYZ GetViewRight(View view) => view.RightDirection.Normalize();
    public static XYZ GetViewUp(View view) => view.UpDirection.Normalize();
    public static XYZ GetViewForward(View view) => view.ViewDirection.Normalize();
}

internal sealed class ElementPickFilter : ISelectionFilter
{
    private readonly Func<Element, bool> _predicate;

    public ElementPickFilter(Func<Element, bool> predicate) => _predicate = predicate;

    public bool AllowElement(Element elem) => _predicate(elem);
    public bool AllowReference(Reference reference, XYZ position) => true;
}
