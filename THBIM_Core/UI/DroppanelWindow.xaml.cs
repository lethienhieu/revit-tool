using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace THBIM.UI
{
    public partial class DroppanelWindow : Window
    {
        private UIDocument _uiDoc;
        private Document _doc;

        public FamilySymbol SelectedFamily { get; private set; }
        public List<Element> SelectedColumns { get; private set; }
        public List<Element> SelectedFloors { get; private set; }
        public bool IsRun { get; private set; } = false;

        public double UserInputLengthMm { get; private set; } = 2500;
        public bool IsLineBasedMode { get; private set; } = true;

        private bool _isPileCapMode = false;

        public DroppanelWindow(UIDocument uiDoc)
        {
            try
            {
                InitializeComponent();

                if (uiDoc == null) throw new ArgumentNullException("UIDocument is null");
                _uiDoc = uiDoc;
                _doc = uiDoc.Document;

                SelectedColumns = new List<Element>();
                SelectedFloors = new List<Element>();

                LoadFamilyTypesGrouped();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khởi tạo Form: {ex.Message}", "THBIM Error");
                this.Close();
            }
        }

        private void Mode_Checked(object sender, RoutedEventArgs e)
        {
            if (_doc == null) return;

            if (rbModePileCap.IsChecked == true)
            {
                _isPileCapMode = true;
                if (grpPlacementMethod != null) grpPlacementMethod.Visibility = System.Windows.Visibility.Collapsed;
            }
            else
            {
                _isPileCapMode = false;
                if (grpPlacementMethod != null) grpPlacementMethod.Visibility = System.Windows.Visibility.Visible;
            }

            LoadFamilyTypesGrouped();
        }

        private void LoadFamilyTypesGrouped()
        {
            try
            {
                if (spFraming == null) return;
                spFraming.Children.Clear();

                BuiltInCategory targetCat;
                string searchKeyword;

                if (_isPileCapMode)
                {
                    targetCat = BuiltInCategory.OST_StructuralFoundation;
                    searchKeyword = "Pilecap";
                }
                else
                {
                    targetCat = BuiltInCategory.OST_StructuralFraming;
                    searchKeyword = "Droppanel";
                }

                var collector = new FilteredElementCollector(_doc)
                    .OfCategory(targetCat)
                    .OfClass(typeof(FamilySymbol))
                    .WhereElementIsElementType()
                    .Cast<FamilySymbol>()
                    .Where(x => x.FamilyName.IndexOf(searchKeyword, StringComparison.OrdinalIgnoreCase) >= 0)
                    .ToList();

                if (collector.Count == 0)
                {
                    Label lbl = new Label
                    {
                        Content = $"No families found containing '{searchKeyword}'!",
                        Foreground = Brushes.Red
                    };
                    spFraming.Children.Add(lbl);
                    return;
                }

                var grouped = collector.GroupBy(x => x.FamilyName).OrderBy(g => g.Key);
                bool isFirst = true;
                foreach (var group in grouped)
                {
                    TextBlock header = new TextBlock
                    {
                        Text = group.Key,
                        FontWeight = FontWeights.Bold,
                        Foreground = Brushes.DarkSlateGray,
                        Margin = new Thickness(0, 10, 0, 5),
                        TextDecorations = TextDecorations.Underline
                    };
                    spFraming.Children.Add(header);

                    foreach (var sym in group)
                    {
                        RadioButton rb = new RadioButton
                        {
                            Content = sym.Name,
                            Tag = sym,
                            Margin = new Thickness(15, 0, 0, 5),
                            GroupName = "DropPanelGroup",
                            IsChecked = isFirst
                        };
                        spFraming.Children.Add(rb);
                        isFirst = false;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi Load Family: " + ex.Message);
            }
        }

        private void chkLineMode_CheckedChanged(object sender, RoutedEventArgs e)
        {
            if (pnlLengthInput == null) return;
            var chk = sender as CheckBox;
            if (chk == null) return;

            if (chk.IsChecked == true)
            {
                pnlLengthInput.Visibility = System.Windows.Visibility.Visible;
                IsLineBasedMode = true;
            }
            else
            {
                pnlLengthInput.Visibility = System.Windows.Visibility.Collapsed;
                IsLineBasedMode = false;
            }
        }

        // --- HÀM PICK & SCAN (LOGIC MỚI: CHỈ ĐẾM GIAO CẮT THỰC) ---
        private void btnPick_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
            try
            {
                ISelectionFilter filter = new ColumnFloorFilter();
                IList<Reference> pickedRefs = _uiDoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, filter, "Chọn Cột và Sàn (Finish khi xong)");

                SelectedColumns.Clear();
                SelectedFloors.Clear();

                foreach (Reference r in pickedRefs)
                {
                    Element elem = _doc.GetElement(r);
                    long catId = elem.Category.Id.GetValue();
                    if (catId == (long)BuiltInCategory.OST_StructuralColumns) SelectedColumns.Add(elem);
                    else if (catId == (long)BuiltInCategory.OST_Floors) SelectedFloors.Add(elem);
                }

                // TÍNH TOÁN GIAO CẮT
                int count = 0;
                View3D view3D = new FilteredElementCollector(_doc)
                    .OfClass(typeof(View3D)).Cast<View3D>()
                    .FirstOrDefault(v => !v.IsTemplate && !v.IsPerspective);

                if (SelectedColumns.Count > 0 && SelectedFloors.Count > 0 && view3D != null)
                {
                    foreach (Element col in SelectedColumns)
                    {
                        BoundingBoxXYZ colBB = col.get_BoundingBox(null);
                        if (colBB == null) continue;
                        XYZ centerBB = (colBB.Min + colBB.Max) / 2.0;

                        // [FIX] Sử dụng Max.Z (Đỉnh cột) thay vì Min.Z (Đáy cột)
                        // Để đảm bảo điểm xuất phát tia nằm ở trên cùng của cột
                        XYZ colPoint = new XYZ(centerBB.X, centerBB.Y, colBB.Max.Z);

                        // 1. Tạo Box tìm kiếm sơ bộ
                        XYZ min = colBB.Min - new XYZ(0.1, 0.1, 0.1);
                        XYZ max = colBB.Max + new XYZ(0.1, 0.1, 0.1);

                        BoundingBoxIntersectsFilter bbFilter = new BoundingBoxIntersectsFilter(new Outline(min, max));
                        var candidateFloors = SelectedFloors.Where(f => bbFilter.PassesFilter(f)).ToList();

                        foreach (var floor in candidateFloors)
                        {
                            // A. Kiểm tra cột có nằm trong vùng biên dạng sàn không (Vertical Check)
                            if (!IsPointOverFloorGeometry(floor, colPoint, view3D)) continue;

                            // B. Kiểm tra cao độ Z (Geometric Intersection Check)
                            double floorTopZ = GetFloorTopZ(floor, colPoint, view3D);
                            if (double.IsNaN(floorTopZ)) continue;

                            double floorThickness = 0;
                            Parameter pThick = floor.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM);
                            if (pThick != null) floorThickness = pThick.AsDouble();
                            double floorBottomZ = floorTopZ - floorThickness;

                            double colBaseZ, colTopZ;
                            GetColumnElevations(col, out colBaseZ, out colTopZ);

                            double tolerance = 0.001;

                            // Logic kiểm tra chân cột
                            bool skipBase = false;
                            if (_isPileCapMode)
                            {
                                if (colBaseZ > floorTopZ + tolerance) skipBase = true;
                            }
                            else
                            {
                                if (colBaseZ >= floorTopZ - tolerance) skipBase = true;
                            }

                            if (skipBase) continue;
                            if (colTopZ <= floorBottomZ + tolerance) continue;

                            count++;
                        }
                    }
                }
                else if (view3D == null)
                {
                    // Fallback
                    foreach (Element col in SelectedColumns)
                    {
                        BoundingBoxXYZ colBB = col.get_BoundingBox(null);
                        if (colBB == null) continue;
                        BoundingBoxIntersectsFilter bbFilter = new BoundingBoxIntersectsFilter(new Outline(colBB.Min, colBB.Max));
                        count += SelectedFloors.Where(f => bbFilter.PassesFilter(f)).Count();
                    }
                }

                if (count > 0)
                {
                    string objectName = _isPileCapMode ? "Pile Caps" : "Drop Panels";
                    txtCount.Text = $"{count} {objectName} Found";
                    txtCount.Foreground = Brushes.Green;
                    btnCreate.IsEnabled = true;
                }
                else
                {
                    txtCount.Text = "0 Intersections";
                    txtCount.Foreground = Brushes.Red;
                    btnCreate.IsEnabled = false;
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException) { }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi tính toán: " + ex.Message);
            }
            finally { this.ShowDialog(); }
        }

        // --- CÁC HÀM HELPER ---

        private double GetFloorTopZ(Element floor, XYZ pt, View3D v)
        {
            try
            {
                ReferenceIntersector ri = new ReferenceIntersector(new List<ElementId> { floor.Id }, FindReferenceTarget.Face, v);
                ri.FindReferencesInRevitLinks = false;

                // [FIX] Bắn tia từ Đỉnh cột + 1 ft (khoảng 300mm) hướng xuống
                // Vì pt bây giờ là Top của cột, nên chỉ cần nhích lên 1 chút là đủ
                var r = ri.FindNearest(new XYZ(pt.X, pt.Y, pt.Z + 1.0), new XYZ(0, 0, -1));

                if (r != null) return r.GetReference().GlobalPoint.Z;
            }
            catch { }
            return double.NaN;
        }

        private void GetColumnElevations(Element col, out double baseZ, out double topZ)
        {
            baseZ = 0;
            topZ = 0;
            try
            {
                Parameter pBaseLvl = col.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM);
                Parameter pBaseOff = col.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM);
                if (pBaseLvl != null)
                {
                    Level l = _doc.GetElement(pBaseLvl.AsElementId()) as Level;
                    baseZ = (l != null ? l.Elevation : 0) + (pBaseOff != null ? pBaseOff.AsDouble() : 0);
                }

                Parameter pTopLvl = col.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM);
                Parameter pTopOff = col.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM);
                if (pTopLvl != null)
                {
                    Level l = _doc.GetElement(pTopLvl.AsElementId()) as Level;
                    topZ = (l != null ? l.Elevation : 0) + (pTopOff != null ? pTopOff.AsDouble() : 0);
                }
            }
            catch { }
        }

        private bool IsPointOverFloorGeometry(Element floor, XYZ colPoint, View3D view3d)
        {
            try
            {
                ReferenceIntersector ri = new ReferenceIntersector(new List<ElementId> { floor.Id }, FindReferenceTarget.Face, view3d);
                ri.FindReferencesInRevitLinks = false;

                // [FIX] Bắn tia từ Đỉnh cột + 1 ft
                return ri.FindNearest(new XYZ(colPoint.X, colPoint.Y, colPoint.Z + 1.0), new XYZ(0, 0, -1)) != null;
            }
            catch { return false; }
        }

        private void btnCreate_Click(object sender, RoutedEventArgs e)
        {
            foreach (var child in spFraming.Children)
            {
                if (child is RadioButton rb && rb.IsChecked == true)
                {
                    SelectedFamily = rb.Tag as FamilySymbol;
                    break;
                }
            }

            if (SelectedFamily == null) { MessageBox.Show("Chưa chọn Family!"); return; }

            if (_isPileCapMode)
            {
                IsLineBasedMode = false;
            }
            else if (IsLineBasedMode)
            {
                if (double.TryParse(txtLength.Text, out double val)) UserInputLengthMm = val;
                else { MessageBox.Show("Chiều dài phải là số (mm)!", "Lỗi"); return; }
            }

            IsRun = true;
            this.Close();
        }
    }

    public class ColumnFloorFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            if (elem.Category == null) return false;
            long catId = elem.Category.Id.GetValue();
            return catId == (long)BuiltInCategory.OST_StructuralColumns || catId == (long)BuiltInCategory.OST_Floors;
        }
        public bool AllowReference(Reference reference, XYZ position) => false;
    }
}