using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace THBIM.Loader;

/// <summary>
/// Base class for all command proxies. Each ribbon button maps to a proxy that
/// forwards execution to the corresponding command in THBIM.Logic.
/// </summary>
public abstract class CommandProxyBase : IExternalCommand
{
    /// <summary>
    /// Key used to look up the real command in Logic's CommandRegistry.
    /// </summary>
    protected abstract string CommandKey { get; }

    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        if (!PluginManager.Instance.IsLoaded)
        {
            message = "THBIM Logic is not loaded. Try clicking the Reload button.";
            return Result.Failed;
        }

        try
        {
            int result = PluginManager.Instance.LogicEntry!.ExecuteCommand(
                CommandKey, commandData, ref message, elements);

            return result switch
            {
                0 => Result.Succeeded,
                1 => Result.Cancelled,
                _ => Result.Failed
            };
        }
        catch (Exception ex)
        {
            message = $"Error executing {CommandKey}: {ex.Message}";
            return Result.Failed;
        }
    }
}
