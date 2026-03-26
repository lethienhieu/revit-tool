using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Windows.Interop; // Cần reference: PresentationCore, WindowsBase
using THBIM.UI;
using THBIM.Licensing;

namespace THBIM
{
    [Transaction(TransactionMode.Manual)]
    public class CallUIQTO : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                if (!THBIM.Licensing.LicenseManager.EnsureActivated(null))
                {
                    return Result.Cancelled;
                }

                if (!THBIM.Licensing.LicenseManager.EnsurePremium())
                    return Result.Cancelled;
                // 1. Khởi tạo Window ngay lập tức
                QTOWindow window = new QTOWindow(commandData);

                // 2. Thiết lập Owner (Để cửa sổ luôn nằm trên Revit, không bị chìm)
                // Yêu cầu thêm: using System.Windows.Interop;
                WindowInteropHelper helper = new WindowInteropHelper(window);
                helper.Owner = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;

                // 3. Hiển thị cửa sổ
                // Dùng Show() để có thể thao tác song song với Revit
                // Hoặc dùng ShowDialog() nếu muốn bắt buộc đóng bảng rồi mới làm việc khác
                window.Show();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = "Lỗi khi mở giao diện: " + ex.Message;
                return Result.Failed;
            }
        }
    }
}