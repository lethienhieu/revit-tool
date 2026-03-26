using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;

namespace THBIM
{
    [Transaction(TransactionMode.Manual)]
    public class AutoDimGrid : IExternalCommand
    {
        // Class lưu dữ liệu hình học
        private class ElementGeometryData
        {
            public Element Element { get; set; }
            public Reference Reference { get; set; }
            public Line Line { get; set; }
            public XYZ Direction { get; set; }
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // 1. Check License (nếu có)
            if (!THBIM.Licensing.LicenseManager.EnsureActivated(null))
            {
                return Result.Cancelled;
            }

            if (!THBIM.Licensing.LicenseManager.EnsurePremium())
                return Result.Cancelled;

            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            View view = doc.ActiveView;

            // ==========================================================
            // BƯỚC 0: HIỆN UI CHỌN STYLE (CHỈ HIỆN 1 LẦN)
            // ==========================================================
            DimensionType userSelectedType = null;

            try
            {
                FormDimStyle form = new FormDimStyle(doc);
                bool? result = form.ShowDialog();

                if (result == true)
                {
                    userSelectedType = form.SelectedDimType;
                }
                else
                {
                    // Nếu bấm Cancel hoặc đóng form thì thoát luôn
                    return Result.Cancelled;
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", "UI Error: " + ex.Message);
                return Result.Failed;
            }

            // ==========================================================
            // BƯỚC 1: VÒNG LẶP LIÊN TỤC (WHILE TRUE)
            // ==========================================================
            while (true)
            {
                try
                {
                    // ------------------------------------------------------
                    // A. CHỌN GRIDS / LEVELS
                    // ------------------------------------------------------
                    // Nếu nhấn ESC ở bước này, Revit sẽ ném ra OperationCanceledException -> Nhảy xuống catch
                    IList<Element> selectedElements = uidoc.Selection.PickElementsByRectangle(new GridLevelSelectionFilter(), "Select Grids or Levels (ESC to exit)");

                    // Nếu không chọn gì mà bấm Finish (trả về list rỗng) -> Tiếp tục vòng lặp hoặc thoát tùy ý
                    if (selectedElements == null || selectedElements.Count == 0) continue;

                    // ------------------------------------------------------
                    // B. XỬ LÝ HÌNH HỌC (GEOMETRY CALCULATION)
                    // ------------------------------------------------------
                    XYZ viewDir = view.ViewDirection;
                    XYZ viewRight = view.RightDirection;

                    List<ElementGeometryData> allData = new List<ElementGeometryData>();

                    foreach (Element elem in selectedElements)
                    {
                        Line geometricLine = null;

                        if (elem is Grid grid)
                        {
                            // Grid trên mặt bằng
                            if (view.ViewType == ViewType.FloorPlan || view.ViewType == ViewType.CeilingPlan || view.ViewType == ViewType.EngineeringPlan)
                            {
                                if (grid.Curve is Line line) geometricLine = line;
                            }
                            // Grid trên mặt đứng
                            else
                            {
                                geometricLine = Line.CreateBound(XYZ.Zero, XYZ.BasisZ);
                            }
                        }
                        else if (elem is Level level)
                        {
                            // Level trên mặt đứng
                            double elev = level.ProjectElevation;
                            XYZ p1 = view.Origin;
                            XYZ p2 = view.Origin + viewRight * 10;
                            p1 = new XYZ(p1.X, p1.Y, elev);
                            p2 = new XYZ(p2.X, p2.Y, elev);
                            geometricLine = Line.CreateBound(p1, p2);
                        }

                        if (geometricLine != null)
                        {
                            allData.Add(new ElementGeometryData
                            {
                                Element = elem,
                                Reference = new Reference(elem),
                                Line = geometricLine,
                                Direction = geometricLine.Direction.Normalize()
                            });
                        }
                    }

                    if (allData.Count < 2)
                    {
                        // Thay vì báo lỗi rồi thoát, ta dùng ShowOverlay hoặc bỏ qua để user chọn lại
                        continue;
                    }

                    // Thuật toán tìm nhóm song song nhiều nhất (Majority Rule)
                    List<List<ElementGeometryData>> groups = new List<List<ElementGeometryData>>();
                    foreach (var item in allData)
                    {
                        bool added = false;
                        foreach (var group in groups)
                        {
                            if (item.Direction.CrossProduct(group[0].Direction).IsZeroLength())
                            {
                                group.Add(item);
                                added = true;
                                break;
                            }
                        }
                        if (!added) groups.Add(new List<ElementGeometryData> { item });
                    }

                    var majorGroup = groups.OrderByDescending(g => g.Count).FirstOrDefault();
                    if (majorGroup == null || majorGroup.Count < 2) continue; // Bỏ qua nếu không tìm thấy nhóm hợp lệ

                    List<Reference> refArray = majorGroup.Select(x => x.Reference).ToList();
                    List<Line> elementLines = majorGroup.Select(x => x.Line).ToList();

                    // ------------------------------------------------------
                    // C. XỬ LÝ WORKPLANE (Để tránh lỗi PickPoint)
                    // ------------------------------------------------------
                    if (view.SketchPlane == null)
                    {
                        using (Transaction t = new Transaction(doc, "Set Temporary WorkPlane"))
                        {
                            t.Start();
                            Plane plane = Plane.CreateByNormalAndOrigin(view.ViewDirection, view.Origin);
                            view.SketchPlane = SketchPlane.Create(doc, plane);
                            doc.Regenerate();
                            t.Commit();
                        }
                    }

                    // ------------------------------------------------------
                    // D. CHỌN ĐIỂM ĐẶT DIM (PICK POINT)
                    // ------------------------------------------------------
                    // Nhấn ESC ở đây cũng sẽ thoát vòng lặp
                    XYZ pickedPoint = uidoc.Selection.PickPoint("Click to place dimension (ESC to exit)");

                    // Tính toán vị trí đường Dim
                    Line closestLine = null;
                    double minDistance = double.MaxValue;
                    XYZ projectionPoint = null;

                    foreach (Line line in elementLines)
                    {
                        XYZ lineDir = line.Direction.Normalize();
                        XYZ v = pickedPoint - line.Origin;
                        double d = v.DotProduct(lineDir);
                        XYZ projected = line.Origin + lineDir * d;
                        double dist = projected.DistanceTo(pickedPoint);

                        if (dist < minDistance)
                        {
                            minDistance = dist;
                            closestLine = line;
                            projectionPoint = projected;
                        }
                    }

                    if (closestLine == null) continue;

                    XYZ elemDir = closestLine.Direction;
                    XYZ dimDir = elemDir.CrossProduct(viewDir).Normalize();

                    // Tạo một đường thẳng ngắn làm placeholder cho vị trí dim
                    XYZ pt1 = projectionPoint - dimDir; // độ dài ảo
                    XYZ pt2 = projectionPoint + dimDir;
                    Line dimLine = Line.CreateBound(pt1, pt2);

                    // ------------------------------------------------------
                    // E. TẠO DIMENSION (TRANSACTION)
                    // ------------------------------------------------------
                    using (Transaction t = new Transaction(doc, "Auto Dim"))
                    {
                        t.Start();

                        ReferenceArray finalRefs = new ReferenceArray();
                        foreach (var r in refArray) finalRefs.Append(r);

                        try
                        {
                            Dimension newDim = doc.Create.NewDimension(view, dimLine, finalRefs);

                            // Áp dụng Style người dùng đã chọn
                            if (userSelectedType != null && newDim != null)
                            {
                                if (newDim.GetTypeId() != userSelectedType.Id)
                                {
                                    newDim.ChangeTypeId(userSelectedType.Id);
                                }
                            }
                        }
                        catch
                        {
                            // Nếu lỗi tạo dim cụ thể này, bỏ qua và cho user chọn lại
                        }

                        t.Commit();
                    }

                    // KẾT THÚC 1 VÒNG LẶP -> QUAY LẠI ĐẦU VÒNG WHILE ĐỂ CHỌN TIẾP
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    // ĐÂY LÀ ĐIỂM THOÁT LỆNH
                    // Khi người dùng nhấn ESC, code sẽ nhảy vào đây
                    break; // Thoát khỏi vòng lặp while(true)
                }
                catch (Exception ex)
                {
                    // Các lỗi khác (không mong muốn)
                    TaskDialog.Show("Error", "Unexpected Error: " + ex.Message);
                    break;
                }
            }

            return Result.Succeeded;
        }
    }

    public class GridLevelSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem) => elem is Grid || elem is Level;
        public bool AllowReference(Reference reference, XYZ position) => false;
    }
}