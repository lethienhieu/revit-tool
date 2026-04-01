namespace Autodesk.Revit.DB
{
    internal static class ElementIdCompat
    {
        internal static long GetValue(this ElementId id)
        {
#if NETFRAMEWORK
            return id.IntegerValue;
#else
            return id.Value;
#endif
        }

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
