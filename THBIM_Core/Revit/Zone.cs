using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace THBIM
{
    public static class Zone
    {
        public static Result Execute(UIDocument uidoc, Document doc, BuiltInCategory bic,
            string zoneP, string zoneV, string numP, string prefix, int start, int digits, ref string msg)
        {
            try
            {
                // 1. Select Region
                Reference refR = uidoc.Selection.PickObject(ObjectType.Element, new FilterRegion(), "Select a Filled Region...");
                FilledRegion region = doc.GetElement(refR) as FilledRegion;
                Solid solid = CreateSolid(region);
                if (solid == null) return Result.Failed;

                // 2. Lấy tất cả phần tử trong vùng
                var allElementsInZone = new FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType()
                    .WherePasses(new ElementIntersectsSolidFilter(solid)).Cast<Element>().ToList();

                if (allElementsInZone.Count == 0) return Result.Cancelled;

                // 3. Lấy Grid ngang
                var horizontalGrids = new FilteredElementCollector(doc).OfClass(typeof(Grid)).Cast<Grid>()
                    .Where(g => IsHorizontal(g))
                    .OrderByDescending(g => g.Curve.GetEndPoint(0).Y).ToList();

                // 4. Chọn nhóm thủ công (Manual Grouping)
                List<List<Element>> manualGroups = new List<List<Element>>();
                HashSet<ElementId> processedIds = new HashSet<ElementId>();

                if (!string.IsNullOrEmpty(numP))
                {
                    int groupIndex = 1;
                    while (processedIds.Count < allElementsInZone.Count)
                    {
                        try
                        {
                            IList<Reference> refs = uidoc.Selection.PickObjects(ObjectType.Element,
                                new ElementIdFilter(allElementsInZone.Select(e => e.Id).ToList()),
                                $"Select Group {groupIndex} (Optional). Click 'Finish' to save. Press ESC to stop grouping.");

                            if (refs.Count > 0)
                            {
                                List<Element> currentGroup = new List<Element>();
                                foreach (var r in refs)
                                {
                                    Element e = doc.GetElement(r);
                                    if (e != null && !processedIds.Contains(e.Id))
                                    {
                                        currentGroup.Add(e);
                                        processedIds.Add(e.Id);
                                    }
                                }
                                if (currentGroup.Count > 0)
                                {
                                    manualGroups.Add(currentGroup);
                                    groupIndex++;
                                }
                            }
                        }
                        catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                        {
                            break; // Dừng chọn nhóm
                        }
                    }
                }

                // 5. TẠO DANH SÁCH BATCH (HÒA TRỘN GROUP VÀ PHẦN TỬ LẺ)
                List<BatchItem> allBatches = new List<BatchItem>();

                // 5a. Thêm các Group đã chọn vào danh sách
                foreach (var group in manualGroups)
                {
                    allBatches.Add(new BatchItem(group, horizontalGrids));
                }

                // 5b. Thêm các phần tử lẻ (chưa được chọn) vào danh sách, mỗi phần tử là 1 batch riêng
                var remainingElements = allElementsInZone.Where(e => !processedIds.Contains(e.Id)).ToList();
                foreach (var el in remainingElements)
                {
                    allBatches.Add(new BatchItem(new List<Element> { el }, horizontalGrids));
                }

                // 6. SẮP XẾP TOÀN CỤC (Sorting Global Batches)
                // Sắp xếp các Batch dựa trên vị trí đại diện của nó (Grid gần nhất -> Trái qua phải)
                var sortedBatches = allBatches
                    .OrderByDescending(b => b.SortGridY) // Ưu tiên Grid cao nhất
                    .ThenBy(b => b.SortCenterX)          // Sau đó Trái qua Phải
                    .ToList();

                // 7. Thực thi Transaction
                using (Transaction t = new Transaction(doc, "Zone & Mixed Numbering"))
                {
                    t.Start();
                    int currentCounter = start;
                    int totalProcessed = 0;

                    foreach (var batch in sortedBatches)
                    {
                        // Lấy danh sách phần tử trong batch và sắp xếp nội bộ (cho chắc chắn)
                        var finalElements = batch.Elements
                            .OrderByDescending(e => GetNearestGridY(e, horizontalGrids))
                            .ThenBy(e => GetPoint(e).X)
                            .ToList();

                        foreach (Element el in finalElements)
                        {
                            if (!string.IsNullOrEmpty(zoneP)) SetVal(el, zoneP, zoneV);

                            if (!string.IsNullOrEmpty(numP))
                            {
                                SetVal(el, numP, prefix + currentCounter.ToString().PadLeft(digits, '0'));
                                currentCounter++;
                            }
                            totalProcessed++;
                        }
                    }
                    t.Commit();
                    TaskDialog.Show("Success", $"Successfully numbered {totalProcessed} elements.");
                }

                return Result.Succeeded;
            }
            catch (Exception ex) { msg = ex.Message; return Result.Failed; }
        }

        // --- Helper Classes & Methods ---

        // Class đại diện cho một "Cục" (có thể là 1 nhóm hoặc 1 phần tử lẻ)
        private class BatchItem
        {
            public List<Element> Elements { get; }
            public double SortGridY { get; }
            public double SortCenterX { get; }

            public BatchItem(List<Element> elements, List<Grid> grids)
            {
                Elements = elements;

                // Tính tâm của cả nhóm
                double sumX = 0, sumY = 0;
                foreach (var e in elements)
                {
                    XYZ p = GetPoint(e);
                    sumX += p.X;
                    sumY += p.Y;
                }
                XYZ center = new XYZ(sumX / elements.Count, sumY / elements.Count, 0);

                // Gán thuộc tính sắp xếp
                SortCenterX = center.X;

                // Tìm Grid gần tâm nhóm nhất
                if (grids.Any())
                {
                    SortGridY = grids.OrderBy(g => Math.Abs(g.Curve.GetEndPoint(0).Y - center.Y))
                                     .First().Curve.GetEndPoint(0).Y;
                }
                else
                {
                    SortGridY = center.Y;
                }
            }
        }

        private static double GetNearestGridY(Element e, List<Grid> hGrids)
        {
            double elementY = GetPoint(e).Y;
            if (!hGrids.Any()) return elementY;
            return hGrids.OrderBy(g => Math.Abs(g.Curve.GetEndPoint(0).Y - elementY))
                         .First().Curve.GetEndPoint(0).Y;
        }

        private static void SetVal(Element e, string p, string v)
        {
            Parameter param = e.LookupParameter(p) ?? e.Document.GetElement(e.GetTypeId())?.LookupParameter(p);
            if (param != null && !param.IsReadOnly) param.Set(v);
        }

        private static XYZ GetPoint(Element e)
        {
            if (e.Location is LocationPoint lp) return lp.Point;
            if (e.Location is LocationCurve lc) return (lc.Curve.GetEndPoint(0) + lc.Curve.GetEndPoint(1)) / 2;
            return e.get_BoundingBox(null)?.Max ?? XYZ.Zero;
        }

        private static bool IsHorizontal(Grid g)
        {
            Line line = g.Curve as Line;
            return line != null && Math.Abs(line.Direction.Y) < 0.01;
        }

        private static Solid CreateSolid(FilledRegion r)
        {
            Options opt = new Options { DetailLevel = ViewDetailLevel.Fine };
            foreach (GeometryObject obj in r.get_Geometry(opt))
            {
                if (obj is Solid s && s.Faces.Size > 0)
                {
                    PlanarFace pf = s.Faces.Cast<Face>().FirstOrDefault(f => f is PlanarFace p && p.FaceNormal.IsAlmostEqualTo(XYZ.BasisZ)) as PlanarFace;
                    if (pf != null)
                    {
                        List<CurveLoop> loops = pf.EdgeLoops.Cast<EdgeArray>().Select(ea => CurveLoop.Create(ea.Cast<Edge>().Select(x => x.AsCurve()).ToList())).ToList();
                        Transform move = Transform.CreateTranslation(new XYZ(0, 0, -500 - pf.Origin.Z));
                        return GeometryCreationUtilities.CreateExtrusionGeometry(loops.Select(l => CurveLoop.CreateViaTransform(l, move)).ToList(), XYZ.BasisZ, 1000);
                    }
                }
            }
            return null;
        }
    }

    public class FilterRegion : ISelectionFilter { public bool AllowElement(Element e) => e is FilledRegion; public bool AllowReference(Reference r, XYZ p) => false; }
    public class ElementIdFilter : ISelectionFilter { private readonly HashSet<ElementId> _ids; public ElementIdFilter(List<ElementId> ids) => _ids = new HashSet<ElementId>(ids); public bool AllowElement(Element e) => _ids.Contains(e.Id); public bool AllowReference(Reference r, XYZ p) => false; }
}