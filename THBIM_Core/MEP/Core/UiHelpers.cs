#if NET8_0_OR_GREATER
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace THBIM.MEP.Core;

internal static class UiHelpers
{
    public static PushButtonData CreateButtonData(CommandMetadata metadata)
    {
        var asmLocation = Assembly.GetExecutingAssembly().Location;
        var button = new PushButtonData(
            metadata.InternalName,
            metadata.DisplayName,
            asmLocation,
            metadata.CommandTypeName)
        {
            ToolTip = metadata.Tooltip,
            LongDescription = metadata.LongDescription ?? metadata.Tooltip
        };

        var iconDir = Path.Combine(Path.GetDirectoryName(asmLocation)!, "Resources", "MEP");
        var iconPath = Path.Combine(iconDir, metadata.IconFile);
        if (File.Exists(iconPath))
        {
            var largeImage = new BitmapImage(new Uri(iconPath));
            button.LargeImage = largeImage;

            // Try to find a 16px version for small image
            var smallName = metadata.IconFile.Replace("32", "16");
            var smallPath = Path.Combine(iconDir, smallName);
            button.Image = File.Exists(smallPath)
                ? new BitmapImage(new Uri(smallPath))
                : largeImage;
        }

        return button;
    }

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
#endif
