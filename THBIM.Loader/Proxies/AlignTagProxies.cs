using Autodesk.Revit.Attributes;

namespace THBIM.Loader.Proxies;

[Transaction(TransactionMode.Manual)]
public class Proxy_AlignTagsLeftCommand : CommandProxyBase
{
    protected override string CommandKey => "AlignTagsLeftCommand";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_AlignTagsRightCommand : CommandProxyBase
{
    protected override string CommandKey => "AlignTagsRightCommand";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_AlignTagsTopCommand : CommandProxyBase
{
    protected override string CommandKey => "AlignTagsTopCommand";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_AlignTagsBottomCommand : CommandProxyBase
{
    protected override string CommandKey => "AlignTagsBottomCommand";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_ArrangeTagsNoCrossCommand : CommandProxyBase
{
    protected override string CommandKey => "ArrangeTagsNoCrossCommand";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_LeaderAngleSettingCommand : CommandProxyBase
{
    protected override string CommandKey => "LeaderAngleSettingCommand";
}
