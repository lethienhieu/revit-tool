using System;
using System.Linq;
using System.Management; // Cần NuGet: System.Management
using System.Security.Cryptography;
using System.Text;
#if NET5_0_OR_GREATER
using System.Runtime.Versioning;
#endif
using Microsoft.Win32;

namespace THBIM.Licensing
{
#if NET5_0_OR_GREATER
    [SupportedOSPlatform("windows")]
#endif
    public static class MachineIdHelper
    {
        private static string _cached;   // cache trong process

        /// <summary>
        /// ID máy: hash SHA256 của chuỗi tổng hợp (UUID/Board/Bios/Disk/Media/Vol/MachineGuid/MachineName) → 16 hex.
        /// </summary>
        public static string Get()
        {
            if (!string.IsNullOrEmpty(_cached)) return _cached;

            try
            {
                // Ưu tiên UUID phần cứng (ổn định trên nhiều máy)
                string uuid = GetWmiOne("Win32_ComputerSystemProduct", "UUID");
                string board = GetWmiOne("Win32_BaseBoard", "SerialNumber");
                string bios = GetWmiOne("Win32_BIOS", "SerialNumber");

                // Một số máy NVMe trả null → gom tất cả ổ
                string disks = string.Join("|", GetWmiMany("Win32_DiskDrive", "SerialNumber"));
                string media = string.Join("|", GetWmiMany("Win32_PhysicalMedia", "SerialNumber"));

                // Volume C:
                string volC = GetWmiOne("Win32_LogicalDisk", "VolumeSerialNumber", "DeviceID='C:'");

                // Registry MachineGuid (là ID cài đặt Windows; dùng làm fallback)
                string machineGuid = ReadReg(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Cryptography", "MachineGuid");

                // Cuối cùng là tên máy (rất yếu, chỉ dùng khi mọi thứ đều rỗng)
                string host = Environment.MachineName;

                var raw = string.Join(";", new[] { uuid, board, bios, disks, media, volC, machineGuid, host }
                                                .Where(s => !string.IsNullOrWhiteSpace(s)));

                using (var sha = SHA256.Create())
                {
                    var bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(raw));
                    _cached = BitConverter.ToString(bytes).Replace("-", "").Substring(0, 16);
                    return _cached;
                }
            }
            catch
            {
                // Fallback cuối
                var s = Environment.MachineName ?? "UNKNOWN";
                using (var sha = SHA256.Create())
                {
                    var bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(s));
                    _cached = BitConverter.ToString(bytes).Replace("-", "").Substring(0, 16);
                    return _cached;
                }
            }
        }

        private static string GetWmiOne(string cls, string prop, string where = null)
        {
            try
            {
                var q = string.IsNullOrWhiteSpace(where) ? $"SELECT {prop} FROM {cls}" : $"SELECT {prop} FROM {cls} WHERE {where}";
                using (var s = new ManagementObjectSearcher(q))
                {
                    foreach (var o in s.Get())
                    {
                        var v = o.Properties[prop]?.Value?.ToString();
                        if (!string.IsNullOrWhiteSpace(v)) return v.Trim();
                    }
                }
            }
            catch { }
            return null;
        }

        private static string[] GetWmiMany(string cls, string prop)
        {
            try
            {
                using (var s = new ManagementObjectSearcher($"SELECT {prop} FROM {cls}"))
                {
                    return s.Get()
                            .Cast<ManagementBaseObject>()
                            .Select(o => o.Properties[prop]?.Value?.ToString())
                            .Where(v => !string.IsNullOrWhiteSpace(v))
                            .Select(v => v.Trim())
                            .ToArray();
                }
            }
            catch { return Array.Empty<string>(); }
        }

        private static string ReadReg(string path, string name)
        {
            try
            {
                // HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Cryptography, MachineGuid
                var v = Registry.GetValue(path, name, null);
                return v?.ToString();
            }
            catch { return null; }
        }
    }
}