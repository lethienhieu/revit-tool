using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.ComponentModel;
using THBIM.Licensing;

namespace THBIM
{
    [Transaction(TransactionMode.Manual)]
    public class CallUIDP : IExternalCommand
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

                // Kiểm tra an toàn: Nếu không có tài liệu đang mở -> Hủy lệnh
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                if (uidoc == null || uidoc.Document == null)
                {
                    message = "Please open a project document first.";
                    return Result.Cancelled;
                }

                // Mở giao diện
                var window = new UI.DroppanelWindow(uidoc);
                window.ShowDialog();

                if (window.IsRun && window.SelectedFamily != null)
                {
                    // Lấy tham số
                    double lengthMm = window.UserInputLengthMm;
                    bool isLineMode = window.IsLineBasedMode;

                    // Gọi Core
                    // UPDATE: Đã thêm window.SelectedFloors vào tham số thứ 3
                    Revit.DroppanelCore.Run(
                        uidoc.Document,
                        window.SelectedColumns,
                        window.SelectedFloors, // <--- QUAN TRỌNG: Truyền danh sách sàn đã chọn sang Core
                        window.SelectedFamily,
                        lengthMm,
                        isLineMode
                    );

                    return Result.Succeeded;
                }
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message + "\n" + ex.StackTrace;
                return Result.Failed;
            }
        }
    }
}