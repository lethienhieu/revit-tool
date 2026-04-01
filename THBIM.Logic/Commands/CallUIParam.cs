using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;




namespace THBIM
{ 




[Transaction(TransactionMode.Manual)]
public class CallUIParam : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        // Mở thẳng giao diện, không cần check user chọn gì cả
        if (!THBIM.Licensing.LicenseManager.EnsureActivated(null))
        {
            return Result.Cancelled;
        }
        if (!THBIM.Licensing.LicenseManager.EnsurePremium())
            return Result.Cancelled;


        Document doc = commandData.Application.ActiveUIDocument.Document;

        CombineParamWindow window = new CombineParamWindow(doc);
        window.ShowDialog();

        return Result.Succeeded;
    }
}
}