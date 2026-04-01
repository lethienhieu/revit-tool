using Autodesk.Revit.Attributes;

namespace THBIM.Loader.Proxies;

[Transaction(TransactionMode.Manual)]
public class Proxy_ProFilterCommand : CommandProxyBase
{
    protected override string CommandKey => "ProFilterCommand";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_ProFilterByFamilyCommand : CommandProxyBase
{
    protected override string CommandKey => "ProFilterByFamilyCommand";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_ProFilterByTypeCommand : CommandProxyBase
{
    protected override string CommandKey => "ProFilterByTypeCommand";
}
