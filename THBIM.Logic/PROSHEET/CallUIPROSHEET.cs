using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows; // Thêm thư viện này để dùng WindowState
using System.Windows.Interop;


namespace THBIM
{
    [Transaction(TransactionMode.Manual)]
    public class CallUIPROSHEET : IExternalCommand
    {
        // Biến tĩnh để giữ cửa sổ không bị mất khi hàm kết thúc
        public static THBIMSheetWindow _window;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            if (!THBIM.Licensing.LicenseManager.EnsureActivated(null))
            {
                return Result.Cancelled;
            }
            if (!THBIM.Licensing.LicenseManager.EnsurePremium())
                return Result.Cancelled;

            // 1. NẾU CỬA SỔ ĐÃ MỞ (VÀ CHƯA BỊ ĐÓNG) -> GỌI NÓ LÊN
            if (_window != null && _window.IsLoaded)
            {
                // Nếu cửa sổ đang bị thu nhỏ dưới Taskbar -> Phóng to nó lên
                if (_window.WindowState == WindowState.Minimized)
                {
                    _window.WindowState = WindowState.Normal;
                }

                //ép cửa sổ nổi bần bật lên trên cùng
                _window.Activate();
                _window.Topmost = true;
                _window.Topmost = false; // Trả về false ngay để không đè vĩnh viễn các ứng dụng khác

                return Result.Succeeded;
            }

            // 2. NẾU CỬA SỔ CHƯA CÓ HOẶC ĐÃ BỊ TẮT (X) -> TẠO MỚI
            RequestHandler handler = new RequestHandler();
            ExternalEvent exEvent = ExternalEvent.Create(handler);

            Autodesk.Revit.DB.Document doc = commandData.Application.ActiveUIDocument.Document;

            _window = new THBIMSheetWindow(doc, exEvent, handler);

            // ==============================================================
            // BƯỚC QUAN TRỌNG NHẤT: Bắt sự kiện khi người dùng tắt Form (dấu X)
            // Phải dọn dẹp biến tĩnh về null để lần sau Revit biết đường tạo lại
            // ==============================================================
            _window.Closed += (s, e) => { _window = null; };

            // Gán cửa sổ Revit làm "Cha" của cửa sổ Tool
            WindowInteropHelper helper = new WindowInteropHelper(_window);

            // Pro Tip: Trên Revit 2019+, dùng cách này lấy MainWindowHandle an toàn và chuẩn xác hơn Process
            helper.Owner = commandData.Application.MainWindowHandle;

            // Show() cho phép thao tác Revit song song (Modeless)
            _window.Show();

            return Result.Succeeded;
        }
    }
}