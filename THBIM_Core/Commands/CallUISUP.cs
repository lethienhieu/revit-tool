using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System;
using HangerGenerator;
using Autodesk.Revit.Attributes;
using THBIM.Licensing;
#nullable disable
namespace THBIM
{
    [Transaction(TransactionMode.Manual)] // ← PHẢI CÓ DÒNG NÀY
    public class CallUISUP : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            if (!THBIM.Licensing.LicenseManager.EnsureActivated(null))
                return Result.Cancelled;
            if (!THBIM.Licensing.LicenseManager.EnsurePremium())
                return Result.Cancelled;

            var window = new SupportGeneratorWindow(commandData, commandData.Application.ActiveUIDocument.Document);
            window.ShowDialog();
            return Result.Succeeded;
        }
    }
}
