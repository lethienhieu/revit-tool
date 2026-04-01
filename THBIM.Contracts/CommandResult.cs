#if NET8_0_OR_GREATER
namespace THBIM.Contracts;

public static class CommandResult
{
    public const int Succeeded = 0;
    public const int Cancelled = 1;
    public const int Failed = -1;
}
#endif
