using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace THBIM.Loader;

/// <summary>
/// Ribbon command to hot-reload THBIM.Logic without restarting Revit.
/// </summary>
[Transaction(TransactionMode.Manual)]
public class ReloadCommand : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        try
        {
            PluginManager.Instance.Reload();

            // Re-initialize Logic with UIControlledApplication is not possible here
            // (we only have ExternalCommandData). Updaters will re-register on next command.

            TaskDialog.Show("THBIM Reload", "THBIM Logic reloaded successfully!");
            return Result.Succeeded;
        }
        catch (Exception ex)
        {
            TaskDialog.Show("THBIM Reload Error",
                $"Failed to reload THBIM Logic:\n\n{ex.Message}\n\n{ex.StackTrace}");
            message = ex.Message;
            return Result.Failed;
        }
    }
}
