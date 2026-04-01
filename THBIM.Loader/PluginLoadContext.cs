using System.Reflection;
using System.Runtime.Loader;

namespace THBIM.Loader;

/// <summary>
/// Custom collectible AssemblyLoadContext for loading THBIM.Logic.
/// Revit API assemblies and shared contracts are NOT loaded here — they resolve from the default context.
/// </summary>
internal sealed class PluginLoadContext : AssemblyLoadContext
{
    private readonly AssemblyDependencyResolver _resolver;

    /// <summary>
    /// Assemblies that must resolve from the default (Revit host) context.
    /// Loading them here would create incompatible type identities.
    /// </summary>
    private static readonly HashSet<string> _defaultContextAssemblies = new(StringComparer.OrdinalIgnoreCase)
    {
        // Revit API
        "RevitAPI", "RevitAPIUI", "RevitAPIIFC", "RevitAPIMacros",
        "AdWindows", "UIFramework", "UIFrameworkServices",

        // Shared contracts (must be same type identity in Loader and Logic)
        "THBIM.Contracts",

        // WPF runtime (already loaded by Revit process)
        "PresentationCore", "PresentationFramework", "WindowsBase", "System.Xaml",
    };

    public PluginLoadContext(string logicDllPath)
        : base(name: "THBIM.Logic", isCollectible: true)
    {
        _resolver = new AssemblyDependencyResolver(logicDllPath);
    }

    protected override Assembly? Load(AssemblyName assemblyName)
    {
        // Let default context handle Revit API, WPF, and shared contracts
        if (assemblyName.Name is not null && _defaultContextAssemblies.Contains(assemblyName.Name))
            return null;

        // Use deps.json resolver for NuGet dependencies (EPPlus, ClosedXML, etc.)
        string? path = _resolver.ResolveAssemblyToPath(assemblyName);
        return path is not null ? LoadFromAssemblyPath(path) : null;
    }

    protected override IntPtr LoadUnmanagedDll(string unmanagedDllName)
    {
        string? path = _resolver.ResolveUnmanagedDllToPath(unmanagedDllName);
        return path is not null ? LoadUnmanagedDllFromPath(path) : IntPtr.Zero;
    }
}
