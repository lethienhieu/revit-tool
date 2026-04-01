using Autodesk.Revit.Attributes;

namespace THBIM.Loader.Proxies;

// TH Tools MEP tab — LT Create
[Transaction(TransactionMode.Manual)]
public class Proxy_BloomCommand : CommandProxyBase
{
    protected override string CommandKey => "BloomCommand";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_TapCommand : CommandProxyBase
{
    protected override string CommandKey => "TapCommand";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_ElbowDown45Command : CommandProxyBase
{
    protected override string CommandKey => "ElbowDown45Command";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_ElbowRightCommand : CommandProxyBase
{
    protected override string CommandKey => "ElbowRightCommand";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_ElbowDownCommand : CommandProxyBase
{
    protected override string CommandKey => "ElbowDownCommand";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_ElbowLeftCommand : CommandProxyBase
{
    protected override string CommandKey => "ElbowLeftCommand";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_ElbowUpCommand : CommandProxyBase
{
    protected override string CommandKey => "ElbowUpCommand";
}

// TH Tools MEP tab — LT Modify
[Transaction(TransactionMode.Manual)]
public class Proxy_FlipMultipleCommand : CommandProxyBase
{
    protected override string CommandKey => "FlipMultipleCommand";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_RotateFitting30Command : CommandProxyBase
{
    protected override string CommandKey => "RotateFitting30Command";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_RotateFitting45Command : CommandProxyBase
{
    protected override string CommandKey => "RotateFitting45Command";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_RotateFitting60Command : CommandProxyBase
{
    protected override string CommandKey => "RotateFitting60Command";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_RotateFitting75Command : CommandProxyBase
{
    protected override string CommandKey => "RotateFitting75Command";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_RotateFitting90Command : CommandProxyBase
{
    protected override string CommandKey => "RotateFitting90Command";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_DeleteSystemCommand : CommandProxyBase
{
    protected override string CommandKey => "DeleteSystemCommand";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_DisconnectCommand : CommandProxyBase
{
    protected override string CommandKey => "DisconnectCommand";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_MoveConnectCommand : CommandProxyBase
{
    protected override string CommandKey => "MoveConnectCommand";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_MoveConnectAlignCommand : CommandProxyBase
{
    protected override string CommandKey => "MoveConnectAlignCommand";
}

// TH Tools MEP tab — LT Align
[Transaction(TransactionMode.Manual)]
public class Proxy_AlignIn3DCommand : CommandProxyBase
{
    protected override string CommandKey => "AlignIn3DCommand";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_AlignBranchCommand : CommandProxyBase
{
    protected override string CommandKey => "AlignBranchCommand";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_AlignBranchPlusCommand : CommandProxyBase
{
    protected override string CommandKey => "AlignBranchPlusCommand";
}

// TH Tools MEP tab — LT Tools
[Transaction(TransactionMode.Manual)]
public class Proxy_SumParameterCommand : CommandProxyBase
{
    protected override string CommandKey => "SumParameterCommand";
}

[Transaction(TransactionMode.Manual)]
public class Proxy_SectionCommand : CommandProxyBase
{
    protected override string CommandKey => "SectionCommand";
}
