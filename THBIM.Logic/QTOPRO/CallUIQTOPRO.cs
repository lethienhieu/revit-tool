using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Diagnostics;
using System.Windows.Interop; // Cần thiết để set Owner cho cửa sổ

namespace THBIM
{
    [Transaction(TransactionMode.Manual)]
    public class CallUIQTOPRO : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            if (!THBIM.Licensing.LicenseManager.EnsureActivated(null))
            {
                return Result.Cancelled;
            }
            if (!THBIM.Licensing.LicenseManager.EnsurePremium())
                return Result.Cancelled;

            try
            {
                // Lấy Document hiện tại
                Document doc = commandData.Application.ActiveUIDocument.Document;

                // Khởi tạo cửa sổ
                MultiCategoryExportWindow window = new MultiCategoryExportWindow(doc);

                // ÉP CỬA SỔ LUÔN NỔI TRÊN NỀN REVIT (Không bị chìm ra sau)
                WindowInteropHelper helper = new WindowInteropHelper(window);
                helper.Owner = Process.GetCurrentProcess().MainWindowHandle;

                // DÙNG SHOW() THAY VÌ SHOWDIALOG() ĐỂ CHO PHÉP TƯƠNG TÁC VỚI REVIT
                window.Show();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}