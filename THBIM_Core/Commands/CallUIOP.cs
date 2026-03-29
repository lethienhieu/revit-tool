using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System;
#nullable disable
using Autodesk.Revit.Attributes;

namespace THBIM
{
    [Transaction(TransactionMode.Manual)]
    public class CallUIOP : IExternalCommand
    {
        private static OPENING.OPENING _openedWindow;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            if (!THBIM.Licensing.LicenseManager.EnsureActivated(null))
                return Result.Cancelled;
            if (!THBIM.Licensing.LicenseManager.EnsurePremium())
                return Result.Cancelled;

            // Singleton: nếu đã mở thì activate lại
            if (_openedWindow != null && _openedWindow.IsLoaded)
            {
                _openedWindow.Activate();
                return Result.Succeeded;
            }

            var window = new OPENING.OPENING(commandData, commandData.Application.ActiveUIDocument.Document);
            _openedWindow = window;
            window.Closed += (s, e) => _openedWindow = null;
            window.ShowDialog();
            return Result.Succeeded;
        }
    }
}
