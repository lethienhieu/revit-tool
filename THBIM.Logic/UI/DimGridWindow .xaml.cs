using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace THBIM
{
    public partial class FormDimStyle : Window
    {
        public DimensionType SelectedDimType { get; private set; }

        public FormDimStyle(Document doc)
        {
            InitializeComponent();
            LoadDimTypes(doc);
        }

        // Kéo thả cửa sổ (vì WindowStyle=None)
        private void Border_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ButtonState == MouseButtonState.Pressed)
            {
                this.DragMove();
            }
        }

        private void LoadDimTypes(Document doc)
        {
            // Lọc các Dimension Type (Chỉ lấy Linear cho Grid/Level)
            List<DimensionType> dimTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(DimensionType))
                .Cast<DimensionType>()
                .Where(x => x.StyleType == DimensionStyleType.Linear)
                .OrderBy(x => x.Name)
                .ToList();

            cmbDimTypes.ItemsSource = dimTypes;

            // Chọn item đầu tiên nếu có
            if (dimTypes.Count > 0)
                cmbDimTypes.SelectedIndex = 0;
        }

        private void btnOk_Click(object sender, RoutedEventArgs e)
        {
            SelectedDimType = cmbDimTypes.SelectedItem as DimensionType;

            if (SelectedDimType == null)
            {
                MessageBox.Show("Please select a dimension type to proceed.", "Selection Missing", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            this.DialogResult = true;
            this.Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }
    }
}