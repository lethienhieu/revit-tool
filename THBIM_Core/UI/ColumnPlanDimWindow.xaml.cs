using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using ColumnAutoDim.Revit;
#nullable disable

namespace THBIM.Tools
{
    public partial class ColumnPlanDimWindow : Window
    {
        private readonly Document _doc;
        private readonly UIDocument _uidoc;

        public ColumnPlanDimWindow(UIDocument uidoc)
        {
            InitializeComponent();
            _uidoc = uidoc;
            _doc = uidoc.Document;

            Loaded += Window_Loaded;
            btnRun.Click += BtnRun_Click;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // 1. Load Preview Images
            if (imgX.Source == null) imgX.Source = LoadPreview("preview_dim_x.png");
            if (imgY.Source == null) imgY.Source = LoadPreview("preview_dim_y.png");
            HookToggleEvents();
            UpdatePreview();

            // 2. Load Dimension Types
            LoadDimTypes();
            LoadTagTypes();
        }

        private void LoadDimTypes()
        {
            var dimTypes = new FilteredElementCollector(_doc)
                .OfClass(typeof(DimensionType))
                .Cast<DimensionType>()
                .Where(dt => dt.StyleType == DimensionStyleType.Linear)
                .OrderBy(dt => dt.Name)
                .ToList();

            cboDimType.ItemsSource = dimTypes;
            cboDimType.DisplayMemberPath = "Name";

            if (dimTypes.Any()) cboDimType.SelectedIndex = 0;
        }
        private void LoadTagTypes()
        {
            var tagTypes = new FilteredElementCollector(_doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_StructuralColumnTags) // Lọc Tag Cột Kết Cấu
                .Cast<FamilySymbol>()
                .OrderBy(t => t.Name)
                .ToList();

            cboTagType.ItemsSource = tagTypes;
            // Nếu DisplayMemberPath="Name" đã khai báo ở XAML thì không cần dòng dưới, nhưng để chắc chắn:
            cboTagType.DisplayMemberPath = "Name";

            if (tagTypes.Any()) cboTagType.SelectedIndex = 0;
        }

        private void BtnRun_Click(object sender, RoutedEventArgs e)
        {
            // --- 1. VALIDATE DATA ---
            double offsetChain = 700.0;
            if (chkAutoChain.IsChecked == false)
            {
                if (!double.TryParse(txtOffsetChain.Text, out offsetChain))
                {
                    MessageBox.Show("Please enter a valid number for Chain Offset.", "Input Error");
                    return;
                }
            }

            double offsetGrid = 800.0;
            if (chkAutoGrid.IsChecked == false)
            {
                if (!double.TryParse(txtOffsetGrid.Text, out offsetGrid))
                {
                    MessageBox.Show("Please enter a valid number for Grid Offset.", "Input Error");
                    return;
                }
            }
            // Lấy giá trị Tag từ ComboBoxItem
            string tagPos = "TR"; // Mặc định
            if (cboTagPos.SelectedItem is System.Windows.Controls.ComboBoxItem item)
            {
                tagPos = item.Tag.ToString();
            }

            // --- CẬP NHẬT SETTINGS ---
            var settings = new DimSettings
            {
                SelectedDimType = cboDimType.SelectedItem as DimensionType,
                // --- LẤY DỮ LIỆU TAG TỪ UI ---
                SelectedTagType = cboTagType.SelectedItem as FamilySymbol,
                IsTagEnabled = chkTag.IsChecked == true,
                TagPosition = tagPos,

                OffsetChainFromColumn = offsetChain,
                OffsetGridStep = offsetGrid,

                // Lấy giá trị từ UI
                IsDimGridEnabled = chkDimGrid.IsChecked == true,
                IsDimColumnEnabled = chkDimCOL.IsChecked == true, // <--- MỚI THÊM

                IsPlaceTop = optTop.IsChecked == true,
                IsPlaceLeft = optLeft.IsChecked == true
            };

            // --- 2. HIDE & SELECT ---
            this.Hide();

            try
            {
                IList<Reference> refs = _uidoc.Selection.PickObjects(
                    ObjectType.Element,
                    new ColumnSelectionFilter(),
                    "Select columns to Dim, then press Finish");

                if (refs.Count > 0)
                {
                    List<Element> cols = refs.Select(r => _doc.GetElement(r)).ToList();
                    List<Grid> grids = new FilteredElementCollector(_doc, _doc.ActiveView.Id)
                        .OfCategory(BuiltInCategory.OST_Grids)
                        .WhereElementIsNotElementType()
                        .Cast<Grid>()
                        .ToList();

                    // --- 3. RUN LOGIC ---
                    AutoDimLogic.Run(_doc, _doc.ActiveView, cols, grids, settings);

                    // --- 4. TÍNH SỐ LƯỢNG ---
                    int processedCount = settings.IsDimColumnEnabled ? cols.Count : 0;

                    // --- 5. FINISH ---
                    this.Close(); // Luôn đóng bảng sau khi chạy xong

                    // CHỈ HIỆN THÔNG BÁO NẾU CÓ CỘT ĐƯỢC XỬ LÝ
                    if (processedCount > 0)
                    {
                        MessageBox.Show($"Successfully created Dimensions for {processedCount} columns!", "THBIM Tools");
                    }
                }
                else
                {
                    this.Show();
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                this.Show();
            }
            catch (Exception ex)
            {
                this.Show();
                MessageBox.Show("Error: " + ex.Message, "Error");
            }
        }


        private void BtnCombine_Click(object sender, RoutedEventArgs e)
        {
            // Ẩn bảng đi để chọn đối tượng
            this.Hide();

            try
            {
                // 1. Chọn các Dim cần nối
                IList<Reference> refs = _uidoc.Selection.PickObjects(
                    ObjectType.Element,
                    new DimensionSelectionFilter(),
                    "Select Dimensions to Combine, then press Finish");

                if (refs.Count > 1)
                {
                    List<Dimension> dimList = refs.Select(r => _doc.GetElement(r) as Dimension).ToList();

                    // 2. Gọi hàm xử lý nối Dim
                    CombineDimLogic.Run(_doc, dimList);
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                // Người dùng nhấn Esc
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
                this.Show();
            }
        }




        // --- IMAGE HELPERS ---
        private static ImageSource LoadPreview(string fileName)
        {
            var asm = Assembly.GetExecutingAssembly();
            var asmName = asm.GetName().Name;
            var baseDir = Path.GetDirectoryName(asm.Location) ?? "";
            var candidates = new[] {
                 $"pack://application:,,,/Resources/{fileName}",
                 $"pack://application:,,,/{asmName};component/Resources/{fileName}",
                 Path.Combine(baseDir, "Resources", fileName)
             };
            foreach (var s in candidates)
            {
                try
                {
                    if (s.StartsWith("pack://")) return new BitmapImage(new Uri(s, UriKind.Absolute));
                    if (File.Exists(s)) return new BitmapImage(new Uri(s, UriKind.Absolute));
                }
                catch { }
            }
            return null;
        }

        private void HookToggleEvents()
        {
            optTop.Checked += (_, __) => UpdatePreview();
            optBottom.Checked += (_, __) => UpdatePreview();
            optLeft.Checked += (_, __) => UpdatePreview();
            optRight.Checked += (_, __) => UpdatePreview();
            optTop.Unchecked += (_, __) => UpdatePreview();
            optBottom.Unchecked += (_, __) => UpdatePreview();
            optLeft.Unchecked += (_, __) => UpdatePreview();
            optRight.Unchecked += (_, __) => UpdatePreview();
        }

        private void UpdatePreview()
        {
            bool showX = (optTop.IsChecked == true) || (optBottom.IsChecked == true);
            imgX.Visibility = showX ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
            imgX.RenderTransformOrigin = new System.Windows.Point(0.5, 0.5);
            imgX.RenderTransform = new RotateTransform(optBottom.IsChecked == true ? 180 : 0);

            bool showY = (optLeft.IsChecked == true) || (optRight.IsChecked == true);
            imgY.Visibility = showY ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
            imgY.RenderTransformOrigin = new System.Windows.Point(0.5, 0.5);
            imgY.RenderTransform = new RotateTransform(optLeft.IsChecked == true ? 0 : 180);
        }
    }

    public class ColumnSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element e)
        {
            if (e.Category == null) return false;
            long catId = e.Category.Id.GetValue();

            // Cho phép chọn Cột Kết cấu, Cột Kiến trúc VÀ Móng Cọc (Foundation)
            return catId == (long)BuiltInCategory.OST_StructuralColumns ||
                   catId == (long)BuiltInCategory.OST_Columns ||
                   catId == (long)BuiltInCategory.OST_StructuralFoundation; // <--- Thêm dòng này
        }
        public bool AllowReference(Reference r, XYZ p) => false;
    }

    public class DimSettings
    {
        public DimensionType SelectedDimType { get; set; }
        public double OffsetChainFromColumn { get; set; }
        public double OffsetGridStep { get; set; }
        public bool IsDimGridEnabled { get; set; }
        public bool IsPlaceTop { get; set; }
        public bool IsPlaceLeft { get; set; }
        public bool IsDimColumnEnabled { get; set; }

        // --- THÊM MỚI CHO TAG ---
        public FamilySymbol SelectedTagType { get; set; } // Loại Tag được chọn
        public bool IsTagEnabled { get; set; }            // Có Tag hay không
        public string TagPosition { get; set; }
    }

    public class DimensionSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element e) => e is Dimension;
        public bool AllowReference(Reference r, XYZ p) => false;
    }
}