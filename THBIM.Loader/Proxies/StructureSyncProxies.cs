using Autodesk.Revit.Attributes;

namespace THBIM.Loader.Proxies;

[Transaction(TransactionMode.Manual)]
public class Proxy_CallUIStructureSync : CommandProxyBase
{
    protected override string CommandKey => "CallUIStructureSync";
}
