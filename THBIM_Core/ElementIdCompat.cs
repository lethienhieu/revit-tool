// Tương thích ElementId API giữa Revit 2023 (net48) và Revit 2025+ (net8)
// net48: IntegerValue (int), ElementId(int)
// net8:  Value (long), ElementId(long)

namespace Autodesk.Revit.DB
{
    internal static class ElementIdCompat
    {
        /// <summary>
        /// Lấy giá trị số của ElementId, tương thích mọi version.
        /// </summary>
        internal static long GetValue(this ElementId id)
        {
#if NETFRAMEWORK
            return id.IntegerValue;
#else
            return id.Value;
#endif
        }

        /// <summary>
        /// Tạo ElementId từ long, tương thích mọi version.
        /// </summary>
        internal static ElementId CreateId(long value)
        {
#if NETFRAMEWORK
            return new ElementId((int)value);
#else
            return new ElementId(value);
#endif
        }
    }
}
