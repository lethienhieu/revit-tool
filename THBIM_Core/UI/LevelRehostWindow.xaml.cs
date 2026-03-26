using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

// Import namespace chứa logic xử lý (Wall.cs, etc.)
using LevelRehost.REVIT;

namespace LevelRehost.UI
{
    public partial class LevelRehostWindow : Window
    {
        private UIDocument _uiDoc;
        private Document _doc;
        private List<Element> _selectedElements; // Danh sách lưu trữ các đối tượng đã chọn

        public LevelRehostWindow(UIDocument uiDoc)
        {
            InitializeComponent();
            _uiDoc = uiDoc;
            _doc = uiDoc.Document;
            _selectedElements = new List<Element>();

            LoadLevels();
        }

        /// <summary>
        /// Load tất cả Level vào ComboBox, sắp xếp theo cao độ
        /// </summary>
        private void LoadLevels()
        {
            List<Level> levels = new FilteredElementCollector(_doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .OrderBy(l => l.ProjectElevation) // Sắp xếp từ thấp lên cao
                .ToList();

            cmbLevels.ItemsSource = levels;
            cmbLevels.DisplayMemberPath = "Name"; // Hiển thị tên Level

            // Tự động chọn Level đầu tiên nếu có
            if (levels.Count > 0) cmbLevels.SelectedIndex = 0;
        }

        /// <summary>
        /// Sự kiện nút "PICK ELEMENTS"
        /// </summary>
        private void BtnSelect_Click(object sender, RoutedEventArgs e)
        {
            // 1. Tạo danh sách Category cho phép dựa trên Checkbox
            List<BuiltInCategory> allowedCats = new List<BuiltInCategory>();
            if (chkWall.IsChecked == true) allowedCats.Add(BuiltInCategory.OST_Walls);
            
            
            if (chkFloor.IsChecked == true) allowedCats.Add(BuiltInCategory.OST_Floors);
            

            
            if (chkDoor.IsChecked == true) allowedCats.Add(BuiltInCategory.OST_Doors);

            if (chkColumn.IsChecked == true) allowedCats.Add(BuiltInCategory.OST_StructuralColumns);
            //if (chkFraming.IsChecked == true) allowedCats.Add(BuiltInCategory.OST_StructuralFraming);

            if (allowedCats.Count == 0)
            {
                MessageBox.Show("Please select at least one Category checkbox.");
                return;
            }

            // 2. Ẩn form để người dùng thao tác
            this.Hide();

            try
            {
                // 3. Gọi lệnh quét chọn với Filter tùy chỉnh
                ISelectionFilter selFilter = new MultiCategorySelectionFilter(allowedCats);

                IList<Reference> refs = _uiDoc.Selection.PickObjects(
                    ObjectType.Element,
                    selFilter,
                    "Select Columns, Beams, Floors or Foundations...");

                // 4. Reset danh sách cũ và thêm danh sách mới
                _selectedElements.Clear();
                foreach (Reference r in refs)
                {
                    Element elem = _doc.GetElement(r);
                    if (elem != null) _selectedElements.Add(elem);
                }

                // 5. Cập nhật UI
                UpdateUIStatus();
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                // Người dùng bấm Esc -> Không làm gì cả
            }
            finally
            {
                // 6. Hiện lại form
                this.ShowDialog();
            }
        }

        /// <summary>
        /// Cập nhật text và màu sắc sau khi chọn
        /// </summary>
        private void UpdateUIStatus()
        {
            int count = _selectedElements.Count;
            txtStatus.Text = $"{count} items";

            if (count > 0)
            {
                // Code mới đã sửa
                txtStatus.Foreground = new SolidColorBrush((System.Windows.Media.Color)ColorConverter.ConvertFromString("#28A745"));
                btnRun.IsEnabled = true;
            }
            else
            {
                txtStatus.Foreground = new SolidColorBrush((System.Windows.Media.Color)ColorConverter.ConvertFromString("#555555"));
                btnRun.IsEnabled = false;
            }
        }

        /// <summary>
        /// Sự kiện nút "APPLY NEW LEVEL"
        /// </summary>
        private void BtnRun_Click(object sender, RoutedEventArgs e)
        {
            // Validation
            if (_selectedElements.Count == 0) return;
            Level targetLevel = cmbLevels.SelectedItem as Level;
            if (targetLevel == null)
            {
                MessageBox.Show("Please select a target Level.");
                return;
            }

            // Thống kê kết quả
            int successCount = 0;
            int failCount = 0;

            // Mở Transaction để thay đổi dữ liệu Revit
            using (Transaction t = new Transaction(_doc, "Rehost Elements"))
            {
                t.Start();

                foreach (Element elem in _selectedElements)
                {
                    bool result = false;

                    // Phân loại đối tượng để gọi hàm xử lý tương ứng
                    // Hiện tại ta mới có WallProcessor, các cái khác sẽ bổ sung sau
                    BuiltInCategory cat = (BuiltInCategory)elem.Category.Id.GetValue();

                    switch (cat)
                    {
                        case BuiltInCategory.OST_Walls: // (Nếu bạn thêm checkbox Wall thì mở cái này)
                                                        // Logic Wall (Ví dụ nếu chọn nhầm tường kiến trúc)
                                                         result = WallProcessor.RehostWall(_doc, elem as Wall, targetLevel);
                            break;

                        case BuiltInCategory.OST_StructuralColumns:
                            result = ColumnProcessor.RehostColumn(_doc, elem as FamilyInstance, targetLevel);
                            break;

                        case BuiltInCategory.OST_StructuralFraming:
                           // result = BeamProcessor.RehostBeam(_doc, elem as FamilyInstance, targetLevel);
                            break;

                        case BuiltInCategory.OST_Floors:
                            result = FloorProcessor.RehostFloor(_doc, elem as Floor, targetLevel);
                            break;

                        case BuiltInCategory.OST_StructuralFoundation:
                            // result = FoundationProcessor.RehostFoundation(_doc, elem as FamilyInstance, targetLevel);
                            break;

                        default:
                            // Thử xử lý Wall ở đây nếu bạn muốn test Wall trước
                            if (elem is Wall wall)
                            {
                                result = WallProcessor.RehostWall(_doc, wall, targetLevel);
                            }
                            break;
                        // Trong vòng lặp foreach, bên trong switch (cat)

                        case BuiltInCategory.OST_Doors:
                            if (elem is FamilyInstance door)
                            {
                                result = DoorProcessor.RehostDoor(_doc, door, targetLevel);
                            }
                            break;
                    }

                    if (result) successCount++;
                    else failCount++;
                }

                t.Commit();
            }

            MessageBox.Show($"Completed.\nSuccess: {successCount}\nFailed/Skipped: {failCount}", "Result");

            // Đóng tool hoặc Reset tùy bạn
            // this.Close(); 
        }
    }

    // --- CLASS FILTER RIÊNG ---
    // Giúp chỉ cho phép chọn đúng Category mong muốn
    public class MultiCategorySelectionFilter : ISelectionFilter
    {
        private readonly List<BuiltInCategory> _allowedCats;

        public MultiCategorySelectionFilter(List<BuiltInCategory> allowedCats)
        {
            _allowedCats = allowedCats;
        }

        public bool AllowElement(Element elem)
        {
            if (elem.Category == null) return false;
            return _allowedCats.Contains((BuiltInCategory)elem.Category.Id.GetValue());
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }
}