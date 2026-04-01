using Autodesk.Revit.Attributes;

namespace THBIM.Loader.Proxies;

[Transaction(TransactionMode.Manual)]
public class Proxy_CallUIDP : CommandProxyBase
{
    protected override string CommandKey => "CallUIDP";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_CallUISplit : CommandProxyBase
{
    protected override string CommandKey => "CallUISplit";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_CallUILV : CommandProxyBase
{
    protected override string CommandKey => "CallUILV";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_AutoPile : CommandProxyBase
{
    protected override string CommandKey => "AutoPile";
}
