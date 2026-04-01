using Autodesk.Revit.Attributes;

namespace THBIM.Loader.Proxies;

[Transaction(TransactionMode.Manual)]
public class Proxy_CallUICol : CommandProxyBase
{
    protected override string CommandKey => "CallUICol";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_CallUIColorslapher : CommandProxyBase
{
    protected override string CommandKey => "CallUIColorslapher";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_CallUIbubble : CommandProxyBase
{
    protected override string CommandKey => "CallUIbubble";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_CallLinkIDs : CommandProxyBase
{
    protected override string CommandKey => "CallLinkIDs";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_CallUIReName : CommandProxyBase
{
    protected override string CommandKey => "CallUIReName";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_CallUIOP : CommandProxyBase
{
    protected override string CommandKey => "CallUIOP";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_CallUIOverlap : CommandProxyBase
{
    protected override string CommandKey => "CallUIOverlap";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_CallUIParam : CommandProxyBase
{
    protected override string CommandKey => "CallUIParam";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_CallUIParaSync : CommandProxyBase
{
    protected override string CommandKey => "CallUIParaSync";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_CallUIFloordrop : CommandProxyBase
{
    protected override string CommandKey => "CallUIFloordrop";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_CallUIZone : CommandProxyBase
{
    protected override string CommandKey => "CallUIZone";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_CallUIPROSHEET : CommandProxyBase
{
    protected override string CommandKey => "CallUIPROSHEET";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_CallUIQTOPRO : CommandProxyBase
{
    protected override string CommandKey => "CallUIQTOPRO";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_SheetLinkCommand : CommandProxyBase
{
    protected override string CommandKey => "SheetLinkCommand";
}
