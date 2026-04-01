using Autodesk.Revit.Attributes;

namespace THBIM.Loader.Proxies;

// MEP panel commands (TH Tools tab)
[Transaction(TransactionMode.Manual)]
public class Proxy_SelectChanged : CommandProxyBase
{
    protected override string CommandKey => "SelectChanged";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_CreateClouds : CommandProxyBase
{
    protected override string CommandKey => "CreateClouds";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_Accept : CommandProxyBase
{
    protected override string CommandKey => "Accept";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_CallSP : CommandProxyBase
{
    protected override string CommandKey => "CallSP";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_CallUIQTO : CommandProxyBase
{
    protected override string CommandKey => "CallUIQTO";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_CallUISUP : CommandProxyBase
{
    protected override string CommandKey => "CallUISUP";
}
