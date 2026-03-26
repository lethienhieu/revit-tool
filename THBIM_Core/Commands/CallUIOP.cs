using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System;
#nullable disable
using Autodesk.Revit.Attributes;

namespace THBIM
{
    [Transaction(TransactionMode.Manual)] // ← PHẢI CÓ DÒNG NÀY
    public class CallUIOP : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            if (!THBIM.Licensing.LicenseManager.EnsureActivated(null))
            {
                return Result.Cancelled;
            }
            if (!THBIM.Licensing.LicenseManager.EnsurePremium())
                return Result.Cancelled;

            var window = new OPENING.OPENING(commandData, commandData.Application.ActiveUIDocument.Document);
            window.ShowDialog();
            return Result.Succeeded;
        }
    }
}
