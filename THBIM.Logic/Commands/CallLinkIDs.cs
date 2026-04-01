#nullable enable

using System;
using System.Collections.Generic;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.Attributes;

namespace THBIM
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CallLinkIDs : IExternalCommand
    {
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



            UIDocument? uidoc = commandData.Application.ActiveUIDocument;
            if (uidoc == null)
            {
                message = "No active document.";
                return Result.Failed;
            }

            Document doc = uidoc.Document;
            List<string> idResults = new();

            try
            {
                while (true)
                {
                    Reference pickedRef = uidoc.Selection.PickObject(
                        ObjectType.LinkedElement,
                        "Select an element (ESC to close)");

                    if (pickedRef.LinkedElementId == ElementId.InvalidElementId)
                        continue;

                    if (doc.GetElement(pickedRef.ElementId) is not RevitLinkInstance linkInstance)
                        continue;

                    Document? linkDoc = linkInstance.GetLinkDocument();
                    if (linkDoc == null)
                        continue;

                    Element? linkedElement = linkDoc.GetElement(pickedRef.LinkedElementId);
                    if (linkedElement == null)
                        continue;

                    ElementId typeId = linkedElement.GetTypeId();
                    Element? typeElem = linkDoc.GetElement(typeId);
                    string typeName = typeElem?.Name ?? "Unknown";

                    idResults.Add(
                        $"[Link Element: {typeName}] ElementId = {pickedRef.LinkedElementId.GetValue()}"
                    );
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                // Người dùng nhấn ESC
                if (idResults.Count == 0)
                {
                    TaskDialog.Show("Result", "No elements were selected.");
                    return Result.Cancelled;
                }

                try
                {
                    LinkedIDSWindow window = new LinkedIDSWindow(idResults);
                    var helper = new System.Windows.Interop.WindowInteropHelper(window)
                    {
                        Owner = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle
                    };
                    window.ShowDialog();
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("error open form", ex.ToString());
                    return Result.Failed;
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Unknown error", ex.ToString());
                return Result.Failed;
            }
        }
    }
}
