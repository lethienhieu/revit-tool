using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

using Document = Autodesk.Revit.DB.Document;
using Transaction = Autodesk.Revit.DB.Transaction;

namespace THBIM
{
    // --- 1. EVENT HANDLER ---
    public class OverlapHandler : IExternalEventHandler
    {
        public OverlapWindow MyWindow { get; set; }
        public BuiltInCategory SelectedCategory { get; set; } = BuiltInCategory.OST_StructuralFoundation;

        public void Execute(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            if (uidoc == null) return;

            List<OverlapGroup> results = CheckOverlapLogic.RunScanAndHighlight(uidoc.Document, SelectedCategory);

            if (MyWindow != null) MyWindow.UpdateResults(results);
        }

        public string GetName() => "Scan Overlap Handler";
    }

    // --- 2. LOGIC CHÍNH ---
    public static class CheckOverlapLogic
    {
        // 100mm = 0.328 feet (Chiều dài giao nhau tối thiểu để báo lỗi)
        private const double LENGTH_THRESHOLD = 0.328084;

        // 3mm = 0.01 feet (Độ dày tối thiểu để xác định không phải là "Chạm da")
        private const double TOUCHING_TOLERANCE = 0.01;

        public static List<CategoryOption> GetAllModelCategories(Document doc)
        {
            Dictionary<long, CategoryOption> foundCategories = new Dictionary<long, CategoryOption>();

            HashSet<long> ignoredBuiltInCategories = new HashSet<long>
            {
                (long)BuiltInCategory.OST_RvtLinks,
                (long)BuiltInCategory.OST_Lines,
                (long)BuiltInCategory.OST_Cameras,
                (long)BuiltInCategory.OST_AnalysisResults,
                (long)BuiltInCategory.OST_SketchLines,
                (long)BuiltInCategory.OST_Rooms,
                (long)BuiltInCategory.OST_Areas,
                (long)BuiltInCategory.OST_Mass,
                (long)BuiltInCategory.OST_ProjectBasePoint,
                (long)BuiltInCategory.OST_SharedBasePoint,
                (long)BuiltInCategory.OST_Site,
                (long)BuiltInCategory.OST_VolumeOfInterest,
                (long)BuiltInCategory.OST_Sheets,
                (long)BuiltInCategory.OST_Schedules,
                (long)BuiltInCategory.OST_RasterImages
            };

            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .WhereElementIsViewIndependent();

            foreach (Element e in collector)
            {
                Category cat = e.Category;
                if (cat == null || cat.CategoryType != CategoryType.Model) continue;
                if (cat.Name.StartsWith("<")) continue;
                if (!cat.IsVisibleInUI) continue;
                if (ignoredBuiltInCategories.Contains(cat.Id.GetValue())) continue;
                if (!cat.HasMaterialQuantities) continue;
                if (e.get_BoundingBox(null) == null) continue;

                long idLong = (long)cat.Id.GetValue();

                if (!foundCategories.ContainsKey(idLong))
                {
                    if (Enum.IsDefined(typeof(BuiltInCategory), idLong))
                    {
                        foundCategories.Add(idLong, new CategoryOption { Name = cat.Name, BIC = (BuiltInCategory)idLong });
                    }
                }
            }
            return foundCategories.Values.OrderBy(x => x.Name).ToList();
        }

        public static List<OverlapGroup> RunScanAndHighlight(Document doc, BuiltInCategory targetCategory)
        {
            List<OverlapGroup> results = new List<OverlapGroup>();
            View activeView = doc.ActiveView;

            if (!activeView.AreGraphicsOverridesAllowed())
            {
                TaskDialog.Show("Error", "View Template is active. Cannot highlight.");
                return results;
            }

            List<Element> elements = new FilteredElementCollector(doc, activeView.Id)
                .OfCategory(targetCategory)
                .WhereElementIsNotElementType()
                .ToElements()
                .ToList();

            if (elements.Count < 2) return results;

            OverrideGraphicSettings redOverride = GetRedOverrideSettings(doc);

            using (Transaction t = new Transaction(doc, "Scan Overlaps"))
            {
                t.Start();

                int groupCount = 1;
                for (int i = 0; i < elements.Count; i++)
                {
                    Element elem1 = elements[i];
                    BoundingBoxXYZ bb1 = elem1.get_BoundingBox(null);
                    if (bb1 == null) continue;

                    for (int j = i + 1; j < elements.Count; j++)
                    {
                        Element elem2 = elements[j];
                        BoundingBoxXYZ bb2 = elem2.get_BoundingBox(null);
                        if (bb2 == null) continue;

                        // 1. KIỂM TRA SƠ BỘ (TOUCHING CHECK)
                        // Kiểm tra xem 2 hộp bao có thực sự "ăn" vào nhau không, hay chỉ chạm nhẹ
                        if (IsTouchingOnly(bb1, bb2)) continue;

                        bool isConflict = false;

                        // 2. NẾU ĐÃ JOIN -> COI LÀ LỖI
                        if (JoinGeometryUtils.AreElementsJoined(doc, elem1, elem2))
                        {
                            isConflict = true;
                        }
                        // 3. NẾU CHƯA JOIN -> KIỂM TRA KHỐI LƯỢNG GIAO NHAU
                        else if (IsReal3DOverlap(elem1, elem2))
                        {
                            isConflict = true;
                        }

                        if (isConflict)
                        {
                            var group = new OverlapGroup($"Conflict #{groupCount}");
                            group.Items.Add(new OverlapItem(elem1));
                            group.Items.Add(new OverlapItem(elem2));
                            results.Add(group);
                            groupCount++;

                            activeView.SetElementOverrides(elem1.Id, redOverride);
                            activeView.SetElementOverrides(elem2.Id, redOverride);
                        }
                    }
                }
                t.Commit();
            }
            return results;
        }

        // --- HÀM MỚI: KIỂM TRA CHẠM MẶT (Top/Bottom/Side) ---
        private static bool IsTouchingOnly(BoundingBoxXYZ bb1, BoundingBoxXYZ bb2)
        {
            // Tính toán vùng giao nhau của 2 hộp bao
            double maxMinX = Math.Max(bb1.Min.X, bb2.Min.X);
            double minMaxX = Math.Min(bb1.Max.X, bb2.Max.X);
            double overlapX = minMaxX - maxMinX;

            double maxMinY = Math.Max(bb1.Min.Y, bb2.Min.Y);
            double minMaxY = Math.Min(bb1.Max.Y, bb2.Max.Y);
            double overlapY = minMaxY - maxMinY;

            double maxMinZ = Math.Max(bb1.Min.Z, bb2.Min.Z);
            double minMaxZ = Math.Min(bb1.Max.Z, bb2.Max.Z);
            double overlapZ = minMaxZ - maxMinZ;

            // Nếu không giao nhau (giá trị âm) -> Bỏ qua
            if (overlapX <= 0 || overlapY <= 0 || overlapZ <= 0) return true;

            // --- QUAN TRỌNG: CHECK ĐỘ DÀY ---
            // Nếu vùng giao nhau ở BẤT KỲ chiều nào nhỏ hơn 3mm -> Coi như chỉ chạm mặt
            // overlapZ < 3mm -> Xử lý trường hợp Cột chồng lên Cột (Top/Bottom)
            // overlapX, Y < 3mm -> Xử lý trường hợp Cột chạm cạnh Cột (Align)
            if (overlapX < TOUCHING_TOLERANCE ||
                overlapY < TOUCHING_TOLERANCE ||
                overlapZ < TOUCHING_TOLERANCE)
            {
                return true; // Chỉ là chạm nhẹ (Touching) -> Bỏ qua
            }

            return false; // Ăn sâu vào nhau -> Tiếp tục kiểm tra kỹ hơn
        }

        private static bool IsReal3DOverlap(Element e1, Element e2)
        {
            ElementIntersectsElementFilter filter = new ElementIntersectsElementFilter(e2);
            if (!filter.PassesFilter(e1)) return false;

            Solid s1 = GetSolid(e1);
            Solid s2 = GetSolid(e2);
            if (s1 == null || s2 == null) return false;

            try
            {
                Solid intersection = BooleanOperationsUtils.ExecuteBooleanOperation(
                    s1, s2, BooleanOperationsType.Intersect);

                if (intersection == null || intersection.Volume <= 0.0001) return false;

                // Kiểm tra lại lần cuối bằng BoundingBox của phần giao khối
                BoundingBoxXYZ bb = intersection.GetBoundingBox();
                if (bb != null)
                {
                    double lenX = bb.Max.X - bb.Min.X;
                    double lenY = bb.Max.Y - bb.Min.Y;
                    double lenZ = bb.Max.Z - bb.Min.Z;

                    // Lại lọc thêm 1 lần nữa cho chắc chắn
                    double minDim = Math.Min(lenX, Math.Min(lenY, lenZ));
                    if (minDim < TOUCHING_TOLERANCE) return false;

                    // Nếu kích thước vùng lỗi đủ lớn (> 100mm) mới báo
                    double maxDim = Math.Max(lenX, Math.Max(lenY, lenZ));
                    if (maxDim > LENGTH_THRESHOLD) return true;
                }
                return false;
            }
            catch
            {
                // Nếu lỗi tính toán hình học (thường do chạm mặt gây lỗi Boolean)
                // Ta trả về False để an toàn (tránh báo lỗi oan)
                return false;
            }
        }

        private static Solid GetSolid(Element e)
        {
            Options opt = new Options { DetailLevel = ViewDetailLevel.Fine, ComputeReferences = true };
            GeometryElement geomElem = e.get_Geometry(opt);
            return GetSolidFromGeometry(geomElem);
        }

        private static Solid GetSolidFromGeometry(GeometryElement geomElem)
        {
            foreach (GeometryObject obj in geomElem)
            {
                if (obj is Solid s && s.Volume > 0) return s;
                if (obj is GeometryInstance gInst)
                {
                    Solid sInst = GetSolidFromGeometry(gInst.GetInstanceGeometry());
                    if (sInst != null) return sInst;
                }
            }
            return null;
        }

        private static OverrideGraphicSettings GetRedOverrideSettings(Document doc)
        {
            FillPatternElement solidFill = new FilteredElementCollector(doc)
                .OfClass(typeof(FillPatternElement))
                .Cast<FillPatternElement>()
                .FirstOrDefault(x => x.GetFillPattern().IsSolidFill);

            OverrideGraphicSettings ogs = new OverrideGraphicSettings();
            Color red = new Color(255, 0, 0);

            if (solidFill != null)
            {
                ogs.SetSurfaceForegroundPatternId(solidFill.Id);
                ogs.SetSurfaceForegroundPatternColor(red);
                ogs.SetCutForegroundPatternId(solidFill.Id);
                ogs.SetCutForegroundPatternColor(red);
            }
            ogs.SetProjectionLineColor(red);
            ogs.SetCutLineColor(red);
            return ogs;
        }
    }

    public class CategoryOption { public string Name { get; set; } public BuiltInCategory BIC { get; set; } }
    public class OverlapGroup { public string GroupName { get; set; } public List<OverlapItem> Items { get; set; } = new List<OverlapItem>(); public OverlapGroup(string name) { GroupName = name; } }
    public class OverlapItem { public string ElementName { get; set; } public ElementId Id { get; set; } public string CategoryName { get; set; } public OverlapItem(Element e) { ElementName = e.Name; Id = e.Id; CategoryName = e.Category?.Name ?? "Unknown"; } }
}