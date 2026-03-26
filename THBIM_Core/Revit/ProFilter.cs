using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace THBIM
{
    [Transaction(TransactionMode.Manual)]
    public class ProFilterCommand : IExternalCommand
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

            try
            {
                // ====================================================
                // STEP 1: PICK SOURCE ELEMENT
                // ====================================================
                Reference pickedRef = null;
                try
                {
                    // Prompt user to pick one element to define the category
                    pickedRef = uidoc.Selection.PickObject(ObjectType.Element, "Step 1: Select a source element to filter by Category");
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    return Result.Cancelled;
                }

                if (pickedRef == null) return Result.Cancelled;

                Element sourceElement = doc.GetElement(pickedRef);
                Category targetCategory = sourceElement.Category;

                if (targetCategory == null)
                {
                    TaskDialog.Show("Error", "The selected element does not have a valid Category.");
                    return Result.Failed;
                }

                // ====================================================
                // STEP 2: SCAN REGION (NO UI)
                // ====================================================

                // Sử dụng tên class mới: ProCategoryFilter
                ProCategoryFilter filter = new ProCategoryFilter(targetCategory.Id);

                try
                {
                    string promptMessage = $"Step 2: Window-select area to pick all [{targetCategory.Name}] elements";

                    IList<Element> pickedElements = uidoc.Selection.PickElementsByRectangle(filter, promptMessage);

                    // ====================================================
                    // STEP 3: SET SELECTION
                    // ====================================================
                    List<ElementId> idsToSelect = pickedElements.Select(e => e.Id).ToList();

                    if (idsToSelect.Count > 0)
                    {
                        uidoc.Selection.SetElementIds(idsToSelect);
                    }
                    else
                    {
                        TaskDialog.Show("ProFilter", $"No {targetCategory.Name} found in the selection region.");
                    }
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    return Result.Cancelled;
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }

    // --- ĐÃ ĐỔI TÊN CLASS ĐỂ TRÁNH LỖI CS0121 ---
    public class ProCategoryFilter : ISelectionFilter
    {
        private ElementId _targetCategoryId;

        public ProCategoryFilter(ElementId targetCategoryId)
        {
            _targetCategoryId = targetCategoryId;
        }

        public bool AllowElement(Element elem)
        {
            if (elem.Category != null && elem.Category.Id == _targetCategoryId)
            {
                return true;
            }
            return false;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }
}