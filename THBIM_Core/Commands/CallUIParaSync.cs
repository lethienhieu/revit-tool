using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace THBIM
{
    [Transaction(TransactionMode.Manual)]
    public class CallUIParaSync : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            if (!THBIM.Licensing.LicenseManager.EnsureActivated(null))
            {
                return Result.Cancelled;
            }
            if (!THBIM.Licensing.LicenseManager.EnsurePremium())
                return Result.Cancelled;


            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            // Sửa lỗi: Truyền uiDoc vào để Processor có thể thực hiện lệnh PickObjects
            ParaSyncWindow window = new ParaSyncWindow(uiDoc);
            window.ShowDialog();

            return Result.Succeeded;
        }
    }
}