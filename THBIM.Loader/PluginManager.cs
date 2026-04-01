using System.IO;
using System.Reflection;
using THBIM.Contracts;

namespace THBIM.Loader;

/// <summary>
/// Manages the lifecycle of THBIM.Logic: Load → Use → Unload → Reload.
/// Singleton accessed by CommandProxyBase and LoaderApp.
/// </summary>
internal sealed class PluginManager
{
    public static PluginManager Instance { get; } = new();

    private PluginLoadContext? _context;
    private WeakReference? _contextRef;
    private string? _currentShadowDir;

    public ILogicEntry? LogicEntry { get; private set; }
    public bool IsLoaded => LogicEntry is not null;

    /// <summary>Base directory containing THBIM.Loader.dll (set by LoaderApp).</summary>
    public string BaseDir { get; set; } = "";

    private string LogicSourceDir => Path.Combine(BaseDir, "Logic");

    private PluginManager() { }

    /// <summary>
    /// Load THBIM.Logic.dll from a shadow copy.
    /// </summary>
    public void Load()
    {
        if (IsLoaded)
            throw new InvalidOperationException("Logic is already loaded. Call Unload() first.");

        // Shadow copy to a NEW timestamped folder (avoids file lock on old folder)
        _currentShadowDir = ShadowCopy();

        string logicDllPath = Path.Combine(_currentShadowDir, "THBIM.Logic.dll");
        if (!File.Exists(logicDllPath))
            throw new FileNotFoundException("THBIM.Logic.dll not found in shadow folder.", logicDllPath);

        // Create collectible context and load Logic assembly
        _context = new PluginLoadContext(logicDllPath);
        _contextRef = new WeakReference(_context);

        Assembly logicAssembly = _context.LoadFromAssemblyPath(logicDllPath);

        // Instantiate the entry point
        Type? entryType = logicAssembly.GetType("THBIM.Logic.LogicEntry");
        if (entryType is null)
            throw new TypeLoadException("Could not find THBIM.Logic.LogicEntry type.");

        LogicEntry = (ILogicEntry)Activator.CreateInstance(entryType)!;
    }

    /// <summary>
    /// Unload the current Logic context and release all references.
    /// </summary>
    public void Unload()
    {
        if (!IsLoaded) return;

        // Step 1: Let Logic clean up (close windows, unregister updaters, dispose events)
        try { LogicEntry!.PrepareForUnload(); } catch { }
        try { LogicEntry!.Dispose(); } catch { }
        LogicEntry = null;

        // Step 2: Unload the context
        _context!.Unload();
        _context = null;

        // Step 3: Force GC to collect (double-collect pattern for finalizers)
        for (int i = 0; i < 3; i++)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        // Step 4: Verify unload
        if (_contextRef is not null && _contextRef.IsAlive)
        {
            System.Diagnostics.Debug.WriteLine(
                "[THBIM.Loader] WARNING: AssemblyLoadContext still alive after unload. " +
                "There may be leaked references preventing GC collection.");
        }
    }

    /// <summary>
    /// Unload current Logic, then load fresh copy.
    /// </summary>
    public void Reload()
    {
        var oldShadow = _currentShadowDir;
        Unload();
        _currentShadowDir = null; // allow old folder to be cleaned
        Load();

        // Try deleting old shadow after new one is loaded
        if (oldShadow is not null)
        {
            try { Directory.Delete(oldShadow, recursive: true); }
            catch { /* still locked, will be cleaned on next startup */ }
        }
    }

    /// <summary>
    /// Called by LoaderApp.OnShutdown. Unload + clean up ALL shadow folders.
    /// </summary>
    public void Shutdown()
    {
        Unload();
        _currentShadowDir = null;
        CleanupOldShadows();
    }

    /// <summary>
    /// Copy all files from Logic/ to a NEW timestamped shadow folder.
    /// Each reload gets a fresh folder — old locked folders are cleaned up best-effort.
    /// Returns the path to the new shadow directory.
    /// </summary>
    private string ShadowCopy()
    {
        if (!Directory.Exists(LogicSourceDir))
            throw new DirectoryNotFoundException($"Logic source directory not found: {LogicSourceDir}");

        // Clean up old shadow folders (best-effort, ignore if locked)
        CleanupOldShadows();

        // Create new timestamped shadow folder
        string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff");
        string shadowDir = Path.Combine(BaseDir, $"LogicShadow_{timestamp}");
        Directory.CreateDirectory(shadowDir);

        // Copy all files recursively
        foreach (string sourceFile in Directory.GetFiles(LogicSourceDir, "*", SearchOption.AllDirectories))
        {
            string relativePath = Path.GetRelativePath(LogicSourceDir, sourceFile);
            string destFile = Path.Combine(shadowDir, relativePath);

            string? destDir = Path.GetDirectoryName(destFile);
            if (destDir is not null)
                Directory.CreateDirectory(destDir);

            File.Copy(sourceFile, destFile, overwrite: false);
        }

        return shadowDir;
    }

    /// <summary>
    /// Try to delete old LogicShadow_* folders. Silently skip any that are still locked.
    /// </summary>
    private void CleanupOldShadows()
    {
        try
        {
            foreach (string dir in Directory.GetDirectories(BaseDir, "LogicShadow*"))
            {
                // Don't delete the folder currently in use
                if (dir == _currentShadowDir) continue;
                try { Directory.Delete(dir, recursive: true); }
                catch { /* still locked by old context, will be cleaned next time */ }
            }
        }
        catch { /* BaseDir listing failed, ignore */ }
    }
}
