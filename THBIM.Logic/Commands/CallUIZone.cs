using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace THBIM
{
    [Transaction(TransactionMode.Manual)]
    public class CallUIZone : IExternalCommand
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
            ZoneWindow window = new ZoneWindow(uidoc.Document);

            if (window.ShowDialog() == true)
            {
                return Zone.Execute(uidoc, uidoc.Document, window.SelectedCategory,
                    window.SelectedZoneParam, window.SelectedZoneValue,
                    window.SelectedNumberParam, window.Prefix,
                    window.StartNumber, window.Digits, ref message);
            }
            return Result.Cancelled;
        }
    }
}