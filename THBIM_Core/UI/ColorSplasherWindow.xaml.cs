using Autodesk.Revit.DB;
using colorslapsher.REVIT;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows; // Dùng cho WPF Window, MessageBox
using System.Windows.Controls;
using System.Windows.Media; // Dùng cho SolidColorBrush

namespace colorslapsher.UI
{
    public partial class ColorSplasherWindow : Window
    {
        private Document _doc;
        private ColorSplasher _coreLogic;
        private List<UiColorItem>? _currentUiItems;
        private Random _rnd = new Random();

        // Helper Class Binding
        public class UiColorItem : INotifyPropertyChanged
        {
            public ColorItem CoreItem { get; set; }

            public string ValueName => CoreItem.ValueName;

            private SolidColorBrush _displayColorBrush;
            public SolidColorBrush DisplayColorBrush
            {
                get { return _displayColorBrush; }
                set
                {
                    _displayColorBrush = value;
                    OnPropertyChanged("DisplayColorBrush");
                }
            }

            public event PropertyChangedEventHandler PropertyChanged;
            protected void OnPropertyChanged(string name)
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
            }
        }

        public ColorSplasherWindow(Document doc)
        {
            InitializeComponent();
            _doc = doc;
            _coreLogic = new ColorSplasher(doc);
            LoadCategories();

            // Lấy Handle (mã định danh) của cửa sổ Revit chính
            IntPtr revitHandle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;

            // Thiết lập cửa sổ này là "con" của Revit -> Không bao giờ bị chìm xuống dưới Revit
            System.Windows.Interop.WindowInteropHelper helper = new System.Windows.Interop.WindowInteropHelper(this);
            helper.Owner = revitHandle;
        }

        private void LoadCategories()
        {
            try
            {
                var cats = _coreLogic.GetCategoriesInActiveView();
                cmbCategories.ItemsSource = cats;
                cmbCategories.DisplayMemberPath = "Name";
                if (cats.Count > 0) cmbCategories.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                // Gọi đích danh System.Windows.MessageBox để tránh nhầm với WinForms
                System.Windows.MessageBox.Show("Error loading categories: " + ex.Message);
            }
        }

        private void cmbCategories_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selectedCat = cmbCategories.SelectedItem as Category;
            if (selectedCat != null)
            {
                var allParams = _coreLogic.GetParametersOfCategory(selectedCat);
                lbParameters.ItemsSource = allParams;
                lbParameters.DisplayMemberPath = "Definition.Name";
                lbValues.ItemsSource = null;
                _currentUiItems = null;
            }
        }

        private void lbParameters_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selectedCat = cmbCategories.SelectedItem as Category;
            var selectedParam = lbParameters.SelectedItem as Parameter;

            if (selectedCat != null && selectedParam != null)
            {
                var coreItems = _coreLogic.AnalyzeValuesAndGenerateColors(selectedCat, selectedParam.Definition.Name);
                _currentUiItems = coreItems.Select(item => new UiColorItem
                {
                    CoreItem = item,
                    DisplayColorBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(item.R, item.G, item.B))
                }).ToList();
                lbValues.ItemsSource = _currentUiItems;
            }
        }

        // --- SỰ KIỆN CLICK MÀU (Dùng Bảng Màu Windows Cũ) ---
        private void ColorBox_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            Border border = sender as Border;
            UiColorItem selectedItem = border.DataContext as UiColorItem;

            if (selectedItem != null)
            {
                // Gọi bảng màu Revit
                using (Autodesk.Revit.UI.ColorSelectionDialog colorDialog = new Autodesk.Revit.UI.ColorSelectionDialog())
                {
                    // Bảng màu sẽ tự động hiện Modal (chặn thao tác bên dưới) và nằm trên cùng
                    Autodesk.Revit.UI.ItemSelectionDialogResult result = colorDialog.Show();

                    if (result == Autodesk.Revit.UI.ItemSelectionDialogResult.Confirmed)
                    {
                        Autodesk.Revit.DB.Color newRevitColor = colorDialog.SelectedColor;

                        // Cập nhật Core
                        selectedItem.CoreItem.RevitColor = newRevitColor;
                        selectedItem.CoreItem.R = newRevitColor.Red;
                        selectedItem.CoreItem.G = newRevitColor.Green;
                        selectedItem.CoreItem.B = newRevitColor.Blue;

                        // Cập nhật UI
                        selectedItem.DisplayColorBrush = new SolidColorBrush(
                            System.Windows.Media.Color.FromRgb(newRevitColor.Red, newRevitColor.Green, newRevitColor.Blue));
                    }
                }
            }
        }

        private void btnRandom_Click(object sender, RoutedEventArgs e)
        {
            if (_currentUiItems == null || _currentUiItems.Count == 0) return;

            // HashSet lưu mã màu đã dùng trong lần random này
            HashSet<int> usedCodes = new HashSet<int>();

            foreach (var item in _currentUiItems)
            {
                // Logic random không trùng (giống bên Core)
                byte r, g, b;
                int code;
                int safety = 0;

                do
                {
                    r = (byte)_rnd.Next(40, 220);
                    g = (byte)_rnd.Next(40, 220);
                    b = (byte)_rnd.Next(40, 220);
                    code = (r << 16) | (g << 8) | b;
                    safety++;
                }
                while (usedCodes.Contains(code) && safety < 500);

                usedCodes.Add(code);

                // Update Core
                item.CoreItem.R = r;
                item.CoreItem.G = g;
                item.CoreItem.B = b;
                item.CoreItem.RevitColor = new Autodesk.Revit.DB.Color(r, g, b);

                // Update UI
                item.DisplayColorBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(r, g, b));
            }

            // Refresh ListBox
            lbValues.ItemsSource = null;
            lbValues.ItemsSource = _currentUiItems;
        }

        private void btnGradient_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.MessageBox.Show("Gradient feature coming soon.", "Info");
        }

        private void btnLegend_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var selectedCat = cmbCategories.SelectedItem as Category;
                var selectedParam = lbParameters.SelectedItem as Parameter;
                if (selectedCat != null && selectedParam != null && _currentUiItems?.Count > 0)
                {
                    var coreList = _currentUiItems.Select(x => x.CoreItem).ToList();
                    _coreLogic.CreateLegendView(selectedCat, selectedParam.Definition.Name, coreList);
                    System.Windows.MessageBox.Show("Legend created!", "Success");
                }
                else System.Windows.MessageBox.Show("Generate colors first.");
            }
            catch (Exception ex) { System.Windows.MessageBox.Show("Error: " + ex.Message); }
        }

        private void btnCreateFilters_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var selectedCat = cmbCategories.SelectedItem as Category;
                var selectedParam = lbParameters.SelectedItem as Parameter;
                if (selectedCat != null && selectedParam != null && _currentUiItems?.Count > 0)
                {
                    FilterPrefixWindow inputDialog = new FilterPrefixWindow();
                    if (inputDialog.ShowDialog() == true)
                    {
                        string prefix = inputDialog.ResultPrefix;
                        var coreList = _currentUiItems.Select(x => x.CoreItem).ToList();
                        _coreLogic.CreateFiltersForValues(selectedCat, selectedParam.Definition.Name, coreList, prefix);
                        System.Windows.MessageBox.Show($"Filters created with prefix '{prefix}_'", "Success");
                    }
                }
                else System.Windows.MessageBox.Show("Select data first.");
            }
            catch (Exception ex) { System.Windows.MessageBox.Show("Error: " + ex.Message); }
        }

        private void btnApply_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var selectedCat = cmbCategories.SelectedItem as Category;
                var selectedParam = lbParameters.SelectedItem as Parameter;
                if (selectedCat != null && selectedParam != null && _currentUiItems != null)
                {
                    var coreList = _currentUiItems.Select(x => x.CoreItem).ToList();
                    _coreLogic.ApplyColorSplash(selectedCat, selectedParam.Definition.Name, coreList);
                    System.Windows.MessageBox.Show("Colors Applied!", "Done");
                }
            }
            catch (Exception ex) { System.Windows.MessageBox.Show("Error: " + ex.Message); }
        }

        private void btnReset_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var selectedCat = cmbCategories.SelectedItem as Category;
                if (selectedCat != null)
                {
                    _coreLogic.ResetGraphics(selectedCat);
                    System.Windows.MessageBox.Show("Reset done for " + selectedCat.Name);
                }
            }
            catch (Exception ex) { System.Windows.MessageBox.Show("Error: " + ex.Message); }
        }
    }
}