using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using LevelRehost.UI; // Gọi Namespace chứa file giao diện Window
using THBIM.Licensing;

namespace THBIM
{
    // TransactionMode.Manual là bắt buộc. 
    // Chúng ta sẽ tự quản lý Transaction bên trong nút "Run" của giao diện UI.
    [Transaction(TransactionMode.Manual)]
    public class CallUILV : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            try
            {

                if (!THBIM.Licensing.LicenseManager.EnsureActivated(null))
                {
                    return Result.Cancelled;
                }
                if (!THBIM.Licensing.LicenseManager.EnsurePremium())
                    return Result.Cancelled;
                // 1. Lấy thông tin về UI và Document hiện hành
                UIApplication uiapp = commandData.Application;
                UIDocument uidoc = uiapp.ActiveUIDocument;

                // Kiểm tra an toàn: Nếu không có tài liệu nào đang mở
                if (uidoc == null)
                {
                    message = "Please open a project first.";
                    return Result.Cancelled;
                }

                // 2. Khởi tạo cửa sổ giao diện (Window)
                // Truyền uidoc vào để UI có thể xử lý việc chọn đối tượng và chạy transaction
                LevelRehostWindow window = new LevelRehostWindow(uidoc);

                // 3. Hiển thị cửa sổ
                // Sử dụng ShowDialog() để mở cửa sổ dạng Modal (Revit sẽ tạm dừng cho đến khi đóng cửa sổ)
                // Đây là cách an toàn nhất cho các tool xử lý dữ liệu đơn giản.
                window.ShowDialog();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                // Bắt lỗi nếu có sự cố bất ngờ khi khởi động tool
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}