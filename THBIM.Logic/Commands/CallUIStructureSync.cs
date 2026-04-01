using System;
using System.Diagnostics; // Dùng cho Process
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using THBIM.Views; // Namespace chứa Window

namespace THBIM
{
    [Transaction(TransactionMode.Manual)]
    public class CallUIStructureSync : IExternalCommand
    {
        // Biến static để giữ cửa sổ không bị mất khi hàm Execute chạy xong
        public static StructureSyncWindow _window;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // 1. Kiểm tra nếu cửa sổ đã mở thì chỉ cần Activate nó lên (tránh mở nhiều cái)
                if (_window != null && _window.IsLoaded)
                {
                    _window.Activate();
                    return Result.Succeeded;
                }

                // 2. Khởi tạo Logic & Window
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                StructureSync logic = new StructureSync(uidoc);

                _window = new StructureSyncWindow(logic);

                // 3. Gán chủ sở hữu (Owner) là Revit để cửa sổ luôn nổi bên trên Revit
                // (Cần reference: System.Windows.Forms và WindowsBase)
                System.Windows.Interop.WindowInteropHelper helper = new System.Windows.Interop.WindowInteropHelper(_window);
                helper.Owner = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;

                // 4. QUAN TRỌNG NHẤT: Dùng .Show() thay vì .ShowDialog()
                _window.Show();

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
