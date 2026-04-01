#if NET8_0_OR_GREATER
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using THBIM.Contracts;

namespace THBIM.Logic;

/// <summary>
/// Entry point for the Logic assembly, called by THBIM.Loader via the ILogicEntry interface.
/// Manages command execution and cleanup for hot-reload.
/// </summary>
public sealed class LogicEntry : ILogicEntry
{
    private bool _disposed;

    public void Initialize(object controlledApplication)
    {
        // Register updaters that need UIControlledApplication
        if (controlledApplication is UIControlledApplication app)
        {
            try { SleeveChangeUpdater.Register(app); } catch { }
        }
    }

    public int ExecuteCommand(string commandKey, object commandData, ref string message, object elements)
    {
        Type? commandType = CommandRegistry.Get(commandKey);
        if (commandType is null)
        {
            message = $"Unknown command: {commandKey}";
            return CommandResult.Failed;
        }

        if (Activator.CreateInstance(commandType) is not IExternalCommand command)
        {
            message = $"Could not create command instance: {commandKey}";
            return CommandResult.Failed;
        }

        var revitCommandData = (ExternalCommandData)commandData;
        var revitElements = (ElementSet)elements;

        Result result = command.Execute(revitCommandData, ref message, revitElements);

        return result switch
        {
            Result.Succeeded => CommandResult.Succeeded,
            Result.Cancelled => CommandResult.Cancelled,
            _ => CommandResult.Failed
        };
    }

    public void PrepareForUnload()
    {
        // 1. Close all open WPF windows owned by this assembly
        CloseAllWindows();

        // 2. Unregister Revit updaters
        try { SleeveChangeUpdater.Unregister(); } catch { }

        // 3. Clear static state that could hold references
        try { ServiceLocator.Reset(); } catch { }
        try { PluginPaths.Reset(); } catch { }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        // Final cleanup if needed
    }

    private static void CloseAllWindows()
    {
        // Close WPF windows that belong to this assembly
        var logicAssembly = typeof(LogicEntry).Assembly;
        var windowsToClose = System.Windows.Application.Current?.Windows
            .OfType<System.Windows.Window>()
            .Where(w => w.GetType().Assembly == logicAssembly)
            .ToList();

        if (windowsToClose is not null)
        {
            foreach (var window in windowsToClose)
            {
                try { window.Close(); } catch { }
            }
        }
    }
}
#endif
