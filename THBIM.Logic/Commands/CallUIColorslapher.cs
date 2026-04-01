using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using THBIM.Licensing;

// Import namespace chứa UI
using colorslapsher.UI;

namespace THBIM
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CallUIColorslapher : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {

                if (!LicenseManager.EnsureActivated())
                {                   
                    return Result.Cancelled;
                }

                if (!THBIM.Licensing.LicenseManager.EnsurePremium())
                    return Result.Cancelled;





                // 1. Lấy thông tin Document hiện hành
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;

                // Kiểm tra an toàn: Nếu không có file nào đang mở
                if (doc == null)
                {
                    message = "No active document found.";
                    return Result.Failed;
                }

                // 2. Khởi tạo cửa sổ giao diện (Window)
                // Truyền biến 'doc' vào để bên UI có thể sử dụng cho các logic xử lý
                ColorSplasherWindow window = new ColorSplasherWindow(doc);

                // 3. Hiển thị cửa sổ
                // Sử dụng ShowDialog() để mở cửa sổ ở chế độ Modal.
                // (Revit sẽ tạm dừng tương tác bên ngoài cho đến khi bạn đóng tool này lại)
                window.ShowDialog();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                // Nếu có lỗi bất ngờ xảy ra lúc khởi động tool, báo lỗi ra cho người dùng
                message = "Failed to launch Color Splasher: " + ex.Message;
                return Result.Failed;
            }
        }
    }
}