using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace THBIM.Tools.UI
{
    public partial class FloordropWindow : Window
    {
        private Document _doc;
        public bool HasUpdated { get; private set; } = false;
        public FamilySymbol SelectedFamilySymbol { get; private set; }
        public bool IsLeftMode { get; private set; }

        public FloordropWindow(Document doc)
        {
            InitializeComponent();
            _doc = doc;
            LoadFamilies();
        }

        private class FamilyOption
        {
            public string Name { get; set; }
            public FamilySymbol Symbol { get; set; }
        }

        private void LoadFamilies()
        {
            // Lấy cả Generic Annotation và Detail Item để chắc chắn không bỏ sót
            var categories = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_GenericAnnotation,
                BuiltInCategory.OST_DetailComponents
            };

            var filter = new ElementMulticategoryFilter(categories);

            var symbols = new FilteredElementCollector(_doc)
                .WherePasses(filter)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .ToList();

            var options = symbols.Select(s => new FamilyOption
            {
                Name = $"{s.FamilyName} : {s.Name}",
                Symbol = s
            })
            .OrderBy(x => x.Name)
            .ToList();

            cmbFamilies.ItemsSource = options;
            if (options.Count > 0) cmbFamilies.SelectedIndex = 0;
        }

        private void BtnApply_Click(object sender, RoutedEventArgs e)
        {
            var selectedOption = cmbFamilies.SelectedItem as FamilyOption;
            if (selectedOption == null)
            {
                TaskDialog.Show("Error", "Please select a family first.");
                return;
            }

            SelectedFamilySymbol = selectedOption.Symbol;
            IsLeftMode = rbLeft.IsChecked == true;

            this.DialogResult = true;
            this.Close();
        }

        // --- UPDATE FUNCTION: USING CATEGORY FILTER + MANUAL CHECK ---
        private void btnUpdate_Click(object sender, RoutedEventArgs e)
        {
            var selectedOption = cmbFamilies.SelectedItem as FamilyOption;
            if (selectedOption == null || selectedOption.Symbol == null)
            {
                TaskDialog.Show("Error", "Please select the family type to update.");
                return;
            }

            var confirmResult = TaskDialog.Show("THBIM",
              $"Please confirm the family you want to update:\n\n" +
              $"👉 {selectedOption.Name}\n\n" +
              "Click OK to proceed.",
                TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.Cancel);

            // Nếu người dùng nhấn Cancel hoặc dấu X -> Thoát hàm, không làm gì cả
            if (confirmResult == TaskDialogResult.Cancel) return;

            var selectedSymbol = selectedOption.Symbol;
            long targetTypeId = selectedSymbol.Id.GetValue();

            // STEP 1: BROAD FILTER BY CATEGORY
            var categories = new List<BuiltInCategory>
    {
        BuiltInCategory.OST_GenericAnnotation,
        BuiltInCategory.OST_DetailComponents
    };
            var filter = new ElementMulticategoryFilter(categories);

            var elementsInView = new FilteredElementCollector(_doc, _doc.ActiveView.Id)
                .WherePasses(filter)
                .WhereElementIsNotElementType()
                .ToElements();

            // STEP 2: MANUAL ITERATION TO FIND THE CORRECT TYPE ID
            var instances = new List<Element>();
            foreach (var elem in elementsInView)
            {
                if (elem.GetTypeId().GetValue() == targetTypeId)
                {
                    instances.Add(elem);
                }
            }

            // Debug: If still not found, report detailed error
            if (instances.Count == 0)
            {
                TaskDialog.Show("THBIM",
                    $"No changes found!"
                    );
                return;
            }

            int successCount = 0;
            int noChangeCount = 0;

            using (Transaction t = new Transaction(_doc, "Update Floor Drops"))
            {
                t.Start();

                foreach (Element inst in instances)
                {
                    LocationPoint loc = inst.Location as LocationPoint;
                    if (loc == null) continue;

                    XYZ point = loc.Point;
                    string newText = GetAdjacentFloorText(_doc, point);

                    if (!string.IsNullOrEmpty(newText))
                    {
                        Parameter param = GetFirstEditableTextParam(inst);
                        if (param != null)
                        {
                            if (param.AsString() != newText)
                            {
                                param.Set(newText);
                                successCount++;
                            }
                            else
                            {
                                noChangeCount++;
                            }
                        }
                    }
                }

                t.Commit();

                // --- QUAN TRỌNG: ĐÁNH DẤU ĐỂ KHÔNG BỊ ROLLBACK ---
                HasUpdated = true;
            }

            // STEP 3: SHOW NOTIFICATION & CLOSE UI
            if (successCount > 0 || noChangeCount > 0)
            {
                TaskDialog.Show("Update Result",
                    $"Update successful!\n" +
                    $"- Values updated: {successCount} families\n" +
                    $"- Values unchanged (already correct): {noChangeCount} families");

                // --- QUAN TRỌNG: ĐÓNG GIAO DIỆN NGAY ---
                this.Close();
            }
            else
            {
                // Trường hợp không tính toán được thì vẫn giữ UI để người dùng kiểm tra lại
                TaskDialog.Show("Update Result", "Families found but could not calculate floor values (floors might not be found below).");
            }
        }
        // =============================================================
        // HELPER CALCULATION METHODS (CORE)
        // =============================================================

        // --- CALCULATION METHOD WITH EXPANDED SCAN AREA (FIX UPDATE ERROR) ---
        private string GetAdjacentFloorText(Document doc, XYZ pickPoint)
        {
            // 1. Bán kính quét (giữ nguyên 3 feet ~ 900mm)
            double radius = 1.0;

            // 2. THAY ĐỔI QUAN TRỌNG: MỞ RỘNG Z VÔ CỰC
            // Thay vì quét xung quanh điểm pick, ta tạo một cột trụ cực cao
            // Bắt đầu từ độ sâu -1000 feet (rất sâu dưới lòng đất)
            XYZ centerBottom = new XYZ(pickPoint.X, pickPoint.Y, -1000.0);

            List<Curve> profile = new List<Curve> {
        Arc.Create(Plane.CreateByNormalAndOrigin(XYZ.BasisZ, centerBottom), radius, 0, Math.PI),
        Arc.Create(Plane.CreateByNormalAndOrigin(XYZ.BasisZ, centerBottom), radius, Math.PI, 2 * Math.PI)
    };

            // Extrude lên cao 2000 feet (tạo thành cột cao ~600 mét cắt qua mọi thứ)
            Solid searchSolid = GeometryCreationUtilities.CreateExtrusionGeometry(new List<CurveLoop> { CurveLoop.Create(profile) }, XYZ.BasisZ, 2000.0);

            // 3. THAY ĐỔI QUAN TRỌNG: QUÉT TOÀN BỘ DOCUMENT (Bỏ ActiveView.Id)
            // Đôi khi sàn nằm ở View Range thấp hơn không hiện trong View này, 
            // nên ta quét toàn bộ sàn trong dự án, sau đó Solid Filter sẽ lọc ra cái nào dính.
            var floors = new FilteredElementCollector(doc) // <--- Bỏ doc.ActiveView.Id
                .OfCategory(BuiltInCategory.OST_Floors)
                .WherePasses(new ElementIntersectsSolidFilter(searchSolid))
                .ToElements();

            if (floors.Count < 2) return null;

            // 4. Lấy 2 sàn CAO NHẤT trong số những sàn tìm được (Logic này đảm bảo lấy đúng sàn ở tầng hiện tại)
            var topFloors = floors.OrderByDescending(x => x.get_BoundingBox(null).Max.Z).Take(2).ToList();

            // Logic sàn dốc
            if (IsFloorSloped(topFloors[0]) || IsFloorSloped(topFloors[1]))
            {
                return "Var.";
            }

            double h1 = topFloors[0].get_BoundingBox(null).Max.Z;
            double h2 = topFloors[1].get_BoundingBox(null).Max.Z;

            double diff = Math.Round(UnitUtils.ConvertFromInternalUnits(Math.Abs(h1 - h2), UnitTypeId.Millimeters), 0);

            if (diff == 0) return null;
            return diff.ToString();
        }

        private bool IsFloorSloped(Element elem)
        {
            Options opt = new Options() { ComputeReferences = true, DetailLevel = ViewDetailLevel.Fine };
            GeometryElement geo = elem.get_Geometry(opt);
            if (geo == null) return false;

            foreach (GeometryObject obj in geo)
            {
                Solid solid = obj as Solid;
                if (solid != null && solid.Volume > 0)
                {
                    foreach (Face face in solid.Faces)
                    {
                        XYZ normal = face.ComputeNormal(new UV(0.5, 0.5));
                        if (normal.Z > 0)
                        {
                            if (Math.Abs(normal.Z) < 0.999) return true;
                        }
                    }
                }
            }
            return false;
        }

        private Parameter GetFirstEditableTextParam(Element e)
        {
            foreach (Parameter p in e.Parameters)
                if (!p.IsReadOnly && p.StorageType == StorageType.String && p.Definition.GetDataType() == SpecTypeId.String.Text) return p;
            return null;
        }
    }
}