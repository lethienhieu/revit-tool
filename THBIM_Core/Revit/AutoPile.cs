using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using DB = Autodesk.Revit.DB;

namespace THBIM
{
    [Transaction(TransactionMode.Manual)]
    public class AutoPile : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, DB.ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            DB.Document doc = uidoc.Document;

            try
            {
                // 1. Mở giao diện cấu hình (Đã loại bỏ phần Zone)
                AutoPileWindow window = new AutoPileWindow(doc);
                if (window.ShowDialog() != true) return Result.Cancelled;

                // 2. Chạy trực tiếp logic tạo cọc (Không cần check IsZoneMode nữa)
                return ExecuteCreateLogic(uidoc, doc, window, ref message);
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        // ====================================================================================
        // LOGIC CREATE PILE: PICK OBJECTS (LOCAL OR LINK)
        // ====================================================================================
        private Result ExecuteCreateLogic(UIDocument uidoc, DB.Document doc, AutoPileWindow window, ref string message)
        {
            try
            {
                // Lấy thông tin từ Window
                DB.FamilySymbol pileSymbol = window.SelectedPileSymbol;
                string paramName = window.OverrideParameterName;
                double? paramValue = window.OverrideValue;
                bool isLinkMode = window.IsLinkMode;

                if (pileSymbol == null) return Result.Failed;

                // Kích hoạt Family Symbol
                using (DB.Transaction tAct = new DB.Transaction(doc, "Activate Symbol"))
                {
                    tAct.Start();
                    if (!pileSymbol.IsActive) pileSymbol.Activate();
                    tAct.Commit();
                }

                // Cấu hình bộ lọc
                PileCapSelectionFilter filter = new PileCapSelectionFilter(doc, isLinkMode);
                int successCount = 0;
                HashSet<DB.ElementId> processedIds = new HashSet<DB.ElementId>();

                using (DB.Transaction t = new DB.Transaction(doc, "Auto Pile Placement"))
                {
                    t.Start();
                    try
                    {
                        IList<DB.Reference> refs;
                        if (isLinkMode)
                            refs = uidoc.Selection.PickObjects(ObjectType.LinkedElement, filter, "LINK MODE: TAB to select Pile Caps...");
                        else
                            refs = uidoc.Selection.PickObjects(ObjectType.Element, filter, "LOCAL MODE: Select Pile Caps...");

                        if (refs == null || refs.Count == 0) return Result.Cancelled;

                        foreach (DB.Reference r in refs)
                        {
                            // Logic chống trùng lặp
                            DB.ElementId checkId = isLinkMode ? r.LinkedElementId : r.ElementId;
                            if (processedIds.Contains(checkId)) continue;
                            processedIds.Add(checkId);

                            DB.XYZ pointGlobal = null;
                            if (isLinkMode)
                            {
                                // Xử lý Link: Lấy tọa độ từ Linked Element và Transform về Local
                                DB.RevitLinkInstance linkInstance = doc.GetElement(r.ElementId) as DB.RevitLinkInstance;
                                DB.Document linkDoc = linkInstance.GetLinkDocument();
                                DB.Element linkedElement = linkDoc.GetElement(r.LinkedElementId);
                                DB.Transform tf = linkInstance.GetTotalTransform();
                                DB.BoundingBoxXYZ bb = linkedElement.get_BoundingBox(null);
                                if (bb == null) continue;
                                DB.XYZ center = (bb.Min + bb.Max) / 2.0;
                                DB.XYZ bottomLocal = new DB.XYZ(center.X, center.Y, bb.Min.Z);
                                pointGlobal = tf.OfPoint(bottomLocal);
                            }
                            else
                            {
                                // Xử lý Local: Lấy tọa độ trực tiếp
                                DB.Element element = doc.GetElement(r);
                                DB.BoundingBoxXYZ bb = element.get_BoundingBox(null);
                                if (bb == null) continue;
                                DB.XYZ center = (bb.Min + bb.Max) / 2.0;
                                pointGlobal = new DB.XYZ(center.X, center.Y, bb.Min.Z);
                            }

                            // Đặt cọc
                            if (pointGlobal != null)
                            {
                                if (PlacePileAtPoint(doc, pointGlobal, pileSymbol, paramName, paramValue))
                                    successCount++;
                            }
                        }
                    }
                    catch (Autodesk.Revit.Exceptions.OperationCanceledException) { return Result.Cancelled; }

                    t.Commit();
                }
                TaskDialog.Show("Success", $"Placed {successCount} piles.");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        // --- CÁC HÀM HỖ TRỢ (HELPER METHODS) ---

        private bool PlacePileAtPoint(DB.Document doc, DB.XYZ pointGlobal, DB.FamilySymbol symbol, string pName, double? pValue)
        {
            try
            {
                DB.Level hostLevel = GetNearestLevel(doc, pointGlobal.Z);
                if (hostLevel == null) return false;

                DB.XYZ insertPoint = new DB.XYZ(pointGlobal.X, pointGlobal.Y, hostLevel.Elevation);

                DB.FamilyInstance pile = doc.Create.NewFamilyInstance(
                    insertPoint,
                    symbol,
                    hostLevel,
                    DB.Structure.StructuralType.Footing);

                double offset = pointGlobal.Z - hostLevel.Elevation;
                SetOffsetParameter(pile, offset);

                if (!string.IsNullOrEmpty(pName) && pValue.HasValue)
                {
                    DB.Parameter pOverride = pile.LookupParameter(pName);
                    if (pOverride == null) pOverride = symbol.LookupParameter(pName);
                    if (pOverride != null && !pOverride.IsReadOnly) pOverride.Set(pValue.Value);
                }
                return true;
            }
            catch { return false; }
        }

        private void SetOffsetParameter(DB.FamilyInstance pile, double value)
        {
            DB.Parameter p = pile.get_Parameter(DB.BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM);
            if (p == null) p = pile.LookupParameter("Height Offset From Level");
            if (p == null) p = pile.LookupParameter("Base Offset");
            if (p != null && !p.IsReadOnly) p.Set(value);
        }

        private DB.Level GetNearestLevel(DB.Document doc, double zCoord)
        {
            return new DB.FilteredElementCollector(doc)
                .OfClass(typeof(DB.Level))
                .Cast<DB.Level>()
                .OrderBy(l => Math.Abs(l.Elevation - zCoord))
                .FirstOrDefault();
        }
    }

    // --- BỘ LỌC CHỌN CỌC/MÓNG (CREATE MODE) ---
    public class PileCapSelectionFilter : ISelectionFilter
    {
        private DB.Document _doc;
        private bool _isLinkMode;

        public PileCapSelectionFilter(DB.Document doc, bool isLinkMode)
        {
            _doc = doc;
            _isLinkMode = isLinkMode;
        }

        public bool AllowElement(DB.Element elem)
        {
            if (_isLinkMode) return elem is DB.RevitLinkInstance;
            return IsValidCategory(elem);
        }

        public bool AllowReference(DB.Reference reference, DB.XYZ position)
        {
            if (_isLinkMode && _doc != null && reference.LinkedElementId != DB.ElementId.InvalidElementId)
            {
                DB.RevitLinkInstance instance = _doc.GetElement(reference.ElementId) as DB.RevitLinkInstance;
                if (instance == null) return false;

                DB.Document linkDoc = instance.GetLinkDocument();
                if (linkDoc == null) return false;

                DB.Element elem = linkDoc.GetElement(reference.LinkedElementId);
                return IsValidCategory(elem);
            }
            return false;
        }

        private bool IsValidCategory(DB.Element elem)
        {
            if (elem == null || elem.Category == null) return false;
            long catId = elem.Category.Id.GetValue();
            return catId == (long)DB.BuiltInCategory.OST_StructuralFoundation ||
                   catId == (long)DB.BuiltInCategory.OST_StructuralFraming ||
                   catId == (long)DB.BuiltInCategory.OST_GenericModel;
        }
    }
}