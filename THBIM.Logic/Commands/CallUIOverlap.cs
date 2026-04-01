using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace THBIM
{
    [Transaction(TransactionMode.Manual)]
    public class CallUIOverlap : IExternalCommand
    {
        public static OverlapWindow _window = null;

        // --- NÂNG CẤP: Dùng Dictionary để lưu riêng từng Category ---
        // Key: BuiltInCategory (ví dụ Cột), Value: Danh sách lỗi của Cột đó
        public static Dictionary<BuiltInCategory, List<OverlapGroup>> GlobalCache
            = new Dictionary<BuiltInCategory, List<OverlapGroup>>();
        // -------------------------------------------------------------

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {

            if (!THBIM.Licensing.LicenseManager.EnsureActivated(null))
            {
                return Result.Cancelled;
            }
            if (!THBIM.Licensing.LicenseManager.EnsurePremium())
                return Result.Cancelled;




            if (_window != null && _window.IsLoaded)
            {
                _window.Activate();
                return Result.Succeeded;
            }

            OverlapHandler handler = new OverlapHandler();
            ExternalEvent exEvent = ExternalEvent.Create(handler);
            UIDocument uidoc = commandData.Application.ActiveUIDocument;

            _window = new OverlapWindow(uidoc, exEvent, handler);
            _window.Show();

            return Result.Succeeded;
        }
    }
}