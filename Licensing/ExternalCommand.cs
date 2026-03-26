using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using THBIM.Licensing;
using System.Windows.Interop; // Cần thêm để dùng WindowInteropHelper

namespace THBIM.Tools
{
    [Transaction(TransactionMode.Manual)]
    public class OpenLicenseCenterCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData c, ref string m, ElementSet s)
        {
            // FIX: Lấy Handle của cửa sổ Revit hiện tại
            var revitWindowHandle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;

            var st = LicenseManager.GetLocalStatus();
            if (!st.IsValid)
            {
                var login = new LoginWindow(); // Màn hình 1

                // FIX: Gán Owner thông qua Helper thay vì Application.Current.MainWindow
                new WindowInteropHelper(login).Owner = revitWindowHandle;

                var ok = login.ShowDialog() == true;    // true khi LOGIN_PWD ok
                if (!ok) return Result.Cancelled;
            }

            var portal = new LicensePortalWindow();   // Màn hình 2

            // FIX: Gán Owner tương tự cho màn hình Portal
            new WindowInteropHelper(portal).Owner = revitWindowHandle;

            portal.ShowDialog();
            return Result.Succeeded;
        }
    }
}