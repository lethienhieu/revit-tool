using System;
using System.Windows.Interop;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using THBIM;
using THBIM.Licensing;
using THBIM.Tools.UI;

namespace THBIM
{
    [Transaction(TransactionMode.Manual)]
    public class CallUIFloordrop : IExternalCommand
    {
        public static FloordropWindow _openedWindow = null;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // 1. Check License
            if (!THBIM.Licensing.LicenseManager.EnsureActivated(null))
            {
                return Result.Cancelled;
            }
            if (!THBIM.Licensing.LicenseManager.EnsurePremium())
                return Result.Cancelled;

            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Singleton Check
            if (_openedWindow != null && _openedWindow.IsLoaded)
            {
                _openedWindow.Activate();
                return Result.Succeeded;
            }

            FloordropWindow window = new FloordropWindow(doc);
            _openedWindow = window;

            WindowInteropHelper helper = new WindowInteropHelper(window);
            helper.Owner = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;

            // Mở giao diện
            bool? dialogResult = window.ShowDialog();

            // Lưu lại trạng thái: Người dùng đã Update chưa?
            bool hasUpdatedData = window.HasUpdated;

            _openedWindow = null;

            // =================================================================
            // TRƯỜNG HỢP 1: NGƯỜI DÙNG TẮT BẢNG (KHÔNG NHẤN START PICK)
            // =================================================================
            if (dialogResult != true)
            {
                // Nếu đã lỡ Update rồi thì phải lưu lại (Succeeded), đừng Cancel để bị Rollback
                if (hasUpdatedData)
                {
                    return Result.Succeeded;
                }
                else
                {
                    return Result.Cancelled;
                }
            }

            // =================================================================
            // TRƯỜNG HỢP 2: NGƯỜI DÙNG NHẤN START PICK
            // =================================================================

            FamilySymbol symbol = window.SelectedFamilySymbol;
            if (symbol == null) return Result.Failed;

            using (Transaction t = new Transaction(doc, "Activate Symbol"))
            {
                t.Start();
                if (!symbol.IsActive) symbol.Activate();
                t.Commit();
            }

            Result pickResult;

            // Chạy logic Pick (Left hoặc Right)
            if (window.IsLeftMode)
            {
                Floordropleft logicLeft = new Floordropleft();
                pickResult = logicLeft.Run(uidoc, symbol, ref message);
            }
            else
            {
                Floordropright logicRight = new Floordropright();
                pickResult = logicRight.Run(uidoc, symbol, ref message);
            }

            // =================================================================
            // QUAN TRỌNG: GHI ĐÈ KẾT QUẢ ĐỂ BẢO VỆ UPDATE
            // =================================================================
            // Nếu người dùng đang Pick dở mà nhấn ESC (pickResult == Cancelled)
            // NHƯNG trước đó họ đã nhấn Update (hasUpdatedData == true)
            // Thì ta bắt buộc phải trả về Succeeded để không bị mất cái Update đó.

            if (pickResult == Result.Cancelled && hasUpdatedData)
            {
                return Result.Succeeded;
            }

            return pickResult;
        }
    }
}