using System;
using System.Windows.Interop;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
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
                return Result.Cancelled;
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

            bool? dialogResult = window.ShowDialog();
            bool hasUpdatedData = window.HasUpdated;
            _openedWindow = null;

            // User closed without clicking START
            if (dialogResult != true)
                return hasUpdatedData ? Result.Succeeded : Result.Cancelled;

            // User clicked START — run picker
            FamilySymbol symbol = window.SelectedFamilySymbol;
            if (symbol == null) return Result.Failed;

            using (Transaction t = new Transaction(doc, "Activate Symbol"))
            {
                t.Start();
                if (!symbol.IsActive) symbol.Activate();
                t.Commit();
            }

            FloorDropPicker picker = new FloorDropPicker();
            Result pickResult = picker.Run(uidoc, symbol, ref message);

            // Protect UPDATE data if user ESC during picking
            if (pickResult == Result.Cancelled && hasUpdatedData)
                return Result.Succeeded;

            return pickResult;
        }
    }
}
