using System.ComponentModel;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using THBIM.Licensing;

namespace THBIM
{
    [Transaction(TransactionMode.Manual)]
    public class CallUIbubble : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            if (!THBIM.Licensing.LicenseManager.EnsureActivated(null))
            {
                return Result.Cancelled;
            }
            if (!THBIM.Licensing.LicenseManager.EnsurePremium())
                return Result.Cancelled;

            // Chỉ hiện UI – modeless; lệnh thực thi được gọi qua ExternalEvent khi bấm radio
            var win = new GridBubble.GridBubbleWindow(commandData);
            new System.Windows.Interop.WindowInteropHelper(win)
            {
                Owner = commandData.Application.MainWindowHandle
            };
            win.Show();   // modeless
            return Result.Succeeded;
        }

    }
}
