using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using THBIM.MEP.Core;

namespace THBIM.MEP.Commands;

[Transaction(TransactionMode.Manual)]
public abstract class CommandBase : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        try
        {
            // License check — all commands require activated + premium
            if (!THBIM.Licensing.LicenseManager.EnsureActivated(null))
                return Result.Cancelled;
            if (!THBIM.Licensing.LicenseManager.EnsurePremium())
                return Result.Cancelled;

            return ExecuteCore(commandData, ref message, elements);
        }
        catch (Autodesk.Revit.Exceptions.OperationCanceledException)
        {
            return Result.Cancelled;
        }
        catch (Exception)
        {
            return Result.Cancelled;
        }
    }

    protected abstract Result ExecuteCore(ExternalCommandData commandData, ref string message, ElementSet elements);

    protected static UIDocument GetUiDoc(ExternalCommandData data) => data.Application.ActiveUIDocument;
    protected static Document GetDoc(ExternalCommandData data) => GetUiDoc(data).Document;
}
