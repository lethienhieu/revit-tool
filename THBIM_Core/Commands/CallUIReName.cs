using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;



namespace THBIM
{
    [Transaction(TransactionMode.Manual)]
    public class CallUIReName : IExternalCommand
    {

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            if (!THBIM.Licensing.LicenseManager.EnsureActivated(null))
            {
                return Result.Cancelled;
            }
            if (!THBIM.Licensing.LicenseManager.EnsurePremium())
                return Result.Cancelled;

            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Lấy danh sách sheet đã chọn
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            if (ids.Count == 0)
            {
                TaskDialog.Show("NOTE", "Please select a sheet before running the command.!");
                return Result.Cancelled;
            }



            List<ViewSheet> sheets = ids
                .Select(id => doc.GetElement(id))
                .OfType<ViewSheet>()
                .ToList();

            if (sheets.Count == 0)
            {
                TaskDialog.Show("NOTE", "No sheet selected.");
                return Result.Cancelled;
            }

            // Hiển thị hộp thoại WPF
            RenameSheetsWindow window = new RenameSheetsWindow(sheets);
            var helper = new System.Windows.Interop.WindowInteropHelper(window);
            helper.Owner = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
            window.ShowDialog();

            // Cửa sổ tự Apply khi bấm OK. Chỉ trả kết quả.
            return (window.DialogResult == true) ? Result.Succeeded : Result.Cancelled;





        }
    }
}
