using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using THBIM.Tools;

namespace THBIM

{
    [Transaction(TransactionMode.Manual)]
    public class CallUICol : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            if (!THBIM.Licensing.LicenseManager.EnsureActivated(null))
            {
                return Result.Cancelled;
            }
            if (!THBIM.Licensing.LicenseManager.EnsurePremium())
                return Result.Cancelled;


            // Mở Window
            var window = new ColumnPlanDimWindow(commandData.Application.ActiveUIDocument);
            window.ShowDialog(); // Dùng ShowDialog để giữ ngữ cảnh (modal)

            return Result.Succeeded;
        }
    }
}