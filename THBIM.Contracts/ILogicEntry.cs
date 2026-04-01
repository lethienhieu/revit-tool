#if NET8_0_OR_GREATER
namespace THBIM.Contracts;

/// <summary>
/// Entry point interface for the Logic assembly.
/// Called by the Loader via the shared Contracts assembly.
/// Uses 'object' for Revit types since Contracts must NOT reference RevitAPI.
/// </summary>
public interface ILogicEntry : IDisposable
{
    /// <summary>
    /// Execute a command by key.
    /// </summary>
    /// <param name="commandKey">Command identifier matching CommandRegistry key</param>
    /// <param name="commandData">ExternalCommandData (cast in Logic)</param>
    /// <param name="message">Ref message string for Revit</param>
    /// <param name="elements">ElementSet (cast in Logic)</param>
    /// <returns>CommandResult value: 0=Succeeded, 1=Cancelled, -1=Failed</returns>
    int ExecuteCommand(string commandKey, object commandData, ref string message, object elements);

    /// <summary>
    /// Prepare for unloading: close WPF windows, unregister updaters, dispose events.
    /// Called before AssemblyLoadContext.Unload().
    /// </summary>
    void PrepareForUnload();

    /// <summary>
    /// Called once after first load to perform any initialization (e.g., register updaters).
    /// </summary>
    /// <param name="controlledApplication">UIControlledApplication as object</param>
    void Initialize(object controlledApplication);
}
#endif
