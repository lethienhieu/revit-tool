using System.IO;
using System.Reflection;

namespace THBIM
{
    internal static class PluginPaths
    {
        public static string BaseDir
        {
            get
            {
                if (_baseDir != null) return _baseDir;

                string asm = Assembly.GetExecutingAssembly().Location;
                if (string.IsNullOrEmpty(asm))
                {
                    _baseDir = string.Empty;
                    return _baseDir;
                }

                string dir = Path.GetDirectoryName(asm) ?? string.Empty;

#if NET8_0_OR_GREATER
                // If loaded from LogicShadow_* or Logic subfolder, go up to "TH Tools"
                string folderName = Path.GetFileName(dir);
                if (folderName.StartsWith("LogicShadow") || folderName == "Logic")
                    dir = Path.GetDirectoryName(dir) ?? dir;
#endif

                _baseDir = dir;
                return _baseDir;
            }
        }

        private static string _baseDir;

        public static void Reset() { _baseDir = null; }
    }
}
