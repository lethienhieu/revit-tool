using Autodesk.Revit.Attributes;

namespace THBIM.Loader.Proxies;

[Transaction(TransactionMode.Manual)]
public class Proxy_AutoDimGrid : CommandProxyBase
{
    protected override string CommandKey => "AutoDimGrid";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_CombineDim : CommandProxyBase
{
    protected override string CommandKey => "CombineDim";
}
