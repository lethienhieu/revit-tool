using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using THBIM.Helpers;

namespace THBIM.UI
{
    public partial class QTOWindow : Window
    {
        private ExternalCommandData _commandData;
        private Document _doc;

        private List<SystemItem>? pipingSystems;
        private List<SystemItem>? ductingSystems;

        private List<Element> currentSelectedElements = new List<Element>();
        private List<Element> selectedElementsByRectangle = new List<Element>();
        private List<Element> _lastSelectedElements = new List<Element>();

        private LengthUnit CurrentLengthUnit = LengthUnit.Millimeter;

        public QTOWindow(ExternalCommandData commandData)
        {
            InitializeComponent(); // ⚠️ Tại đây Radio_Checked được gọi, nhưng _doc đang null -> Gây lỗi

            _commandData = commandData;
            _doc = commandData.Application.ActiveUIDocument.Document; // ✅ _doc được gán ở đây

            cbLengthUnit.SelectedIndex = 0;
            LoadSystemTypes();
            var wa = SystemParameters.WorkArea;

            this.MaxWidth = wa.Width * 0.95;
            this.MaxHeight = wa.Height * 0.95;

            if (this.Width > this.MaxWidth) this.Width = this.MaxWidth;
            if (this.Height > this.MaxHeight) this.Height = this.MaxHeight;
        }

        private void LoadSystemTypes()
        {
            try
            {
                pipingSystems = QTOHelper.GetUsedSystemTypeNames(_doc)
                    .Select(name => new SystemItem { SystemName = name, IsSelected = false }).ToList();

                ductingSystems = QTOHelper.GetUsedDuctSystemTypeNames(_doc)
                    .Select(name => new SystemItem { SystemName = name, IsSelected = false }).ToList();

                SystemGrid.ItemsSource = pipingSystems;
                PipeGroup.Visibility = System.Windows.Visibility.Visible;
                DuctGroup.Visibility = System.Windows.Visibility.Collapsed;

                OnThongKeClick(null, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading systems: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Radio_Checked(object sender, RoutedEventArgs e)
        {
            // ✅ FIX LỖI: Nếu _doc chưa khởi tạo xong thì không làm gì cả
            if (_doc == null) return;

            if (SystemGrid == null || PipeGroup == null || DuctGroup == null)
                return;

            if (rbPiping.IsChecked == true)
            {
                SystemGrid.ItemsSource = pipingSystems;
                PipeGroup.Visibility = System.Windows.Visibility.Visible;
                DuctGroup.Visibility = System.Windows.Visibility.Collapsed;
            }
            else if (rbDucting.IsChecked == true)
            {
                SystemGrid.ItemsSource = ductingSystems;
                PipeGroup.Visibility = System.Windows.Visibility.Collapsed;
                DuctGroup.Visibility = System.Windows.Visibility.Visible;
            }

            OnThongKeClick(null, null);
        }

        private void LengthUnit_Changed(object sender, SelectionChangedEventArgs e)
        {
            // ✅ FIX LỖI: Kiểm tra an toàn
            if (_doc == null) return;

            if (cbLengthUnit.SelectedItem is ComboBoxItem selectedItem)
            {
                string unit = selectedItem.Content.ToString() ?? "mm";

                switch (unit)
                {
                    case "m":
                        CurrentLengthUnit = LengthUnit.Meter;
                        break;
                    case "inch":
                        CurrentLengthUnit = LengthUnit.Inch;
                        break;
                    default:
                        CurrentLengthUnit = LengthUnit.Millimeter;
                        break;
                }

                if (_lastSelectedElements != null && _lastSelectedElements.Any())
                {
                    LoadFromElements(_doc, _lastSelectedElements, CurrentLengthUnit, this, _doc.ActiveView);
                }
                else
                {
                    OnThongKeClick(null, null);
                }
            }
        }
        private void OnSystemCheckBoxClick(object sender, RoutedEventArgs e)
        {
            // 1. Lấy đối tượng CheckBox và Item tương ứng
            if (sender is CheckBox checkBox && checkBox.DataContext is SystemItem clickedItem)
            {
                bool newState = checkBox.IsChecked == true;

                // 2. Kiểm tra xem người dùng có đang chọn nhiều dòng (SelectedItems) không?
                // Và dòng vừa bấm tích có nằm trong số dòng đang được chọn đó không?
                if (SystemGrid.SelectedItems.Count > 1 && SystemGrid.SelectedItems.Contains(clickedItem))
                {
                    // Duyệt qua tất cả các dòng đang được bôi xanh
                    foreach (var item in SystemGrid.SelectedItems)
                    {
                        if (item is SystemItem systemItem)
                        {
                            systemItem.IsSelected = newState;
                        }
                    }

                    // Cập nhật lại giao diện để các ô khác cũng hiển thị tích
                    SystemGrid.Items.Refresh();
                }

                // 3. Tính toán lại thống kê ngay lập tức
                OnThongKeClick(null, null);
            }
        }

        private void OnThongKeClick(object? sender, RoutedEventArgs? e)
        {
            // ✅ FIX LỖI: Kiểm tra an toàn
            if (_doc == null) return;

            try
            {
                if (rbPiping.IsChecked == true)
                {
                    var selected = pipingSystems?
                        .Where(x => x.IsSelected)
                        .Select(x => x.SystemName)
                        .ToList() ?? new List<string>();

                    PipeGrid.ItemsSource = QTOHelper.GetPipeInfo(_doc, selected, CurrentLengthUnit);
                    FittingGrid.ItemsSource = QTOHelper.GetFittingInfo(_doc, selected);
                    InsulationGrid.ItemsSource = QTOHelper.GetInsulationInfo(_doc, selected, CurrentLengthUnit);
                }
                else if (rbDucting.IsChecked == true)
                {
                    var selected = ductingSystems?
                        .Where(x => x.IsSelected)
                        .Select(x => x.SystemName)
                        .ToList() ?? new List<string>();

                    DuctGrid.ItemsSource = QTOHelper.GetDuctInfo(_doc, selected, CurrentLengthUnit);

#pragma warning disable CS0612
                    FittingGrid.ItemsSource = QTOHelper.GetDuctFittingInfo(_doc, selected);
#pragma warning restore CS0612

                    InsulationGrid.ItemsSource = QTOHelper.GetDuctInsulationInfo(_doc, selected, CurrentLengthUnit);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error during calculation: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ... (Các hàm còn lại giữ nguyên như cũ) ...

        public void SetSelectedElementsFromRectangle(List<Element> elements)
        {
            selectedElementsByRectangle = elements;
        }

        private void OnExportExcelClick(object sender, RoutedEventArgs e)
        {
            var saveFileDialog = new SaveFileDialog
            {
                // 👇 ĐỔI ĐUÔI FILE THÀNH .CSV
                Filter = "CSV File|*.csv|Text File|*.txt",
                FileName = "QuickMechanicalQTO_Export.csv"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                string filePath = saveFileDialog.FileName;
                try
                {
                    // (Đoạn code thu thập dữ liệu dataMap giữ nguyên như cũ)
                    var dataMap = new Dictionary<string, IEnumerable<object>>();
                    void AddSheet(string sheetName, System.Collections.IEnumerable source)
                    {
                        if (source != null)
                        {
                            var list = source.Cast<object>().ToList();
                            if (list.Any()) dataMap[sheetName] = list;
                        }
                    }

                    if (rbPiping.IsChecked == true)
                    {
                        AddSheet("Pipes", PipeGrid.ItemsSource);
                        AddSheet("Fittings", FittingGrid.ItemsSource);
                        AddSheet("Insulation", InsulationGrid.ItemsSource);
                    }
                    else if (rbDucting.IsChecked == true)
                    {
                        AddSheet("Ducts", DuctGrid.ItemsSource);
                        AddSheet("Fittings", FittingGrid.ItemsSource);
                        AddSheet("Insulation", InsulationGrid.ItemsSource);
                    }

                    if (dataMap.Count == 0)
                    {
                        MessageBox.Show("No data to export.", "Notification", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    // 👇 GỌI HÀM MỚI (ExcelExporter.ExportToCsv)
                    ExcelExporter.ExportToCsv(dataMap, filePath);

                    MessageBox.Show("Export successful! You can open the file with Excel.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);

                    // Tự động mở file lên sau khi xuất xong (Tiện lợi)
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error exporting:\n" + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
        public void LoadFromElements(Document doc, List<Element> elements, LengthUnit unit, QTOWindow window, View view)
        {
            if (elements == null || elements.Count == 0) return;
            try
            {
                bool hasPipe = elements.Any(e => e is Pipe);
                bool hasDuct = elements.Any(e => e is Duct);
                var allFittings = new List<QTO_Fitting>();
                var allInsulations = new List<QTO_Insulation>();

                if (hasPipe)
                {
                    rbPiping.IsChecked = true;
                    SystemGrid.Visibility = System.Windows.Visibility.Collapsed;
                    PipeGroup.Visibility = System.Windows.Visibility.Visible;
                    PipeGrid.ItemsSource = QTOHelper.GetPipeInfoFromElements(doc, elements, unit);
                    allFittings.AddRange(QTOHelper.GetFittingInfoFromElements(doc, elements));
                    allInsulations.AddRange(QTOHelper.GetInsulationInfoFromElements(doc, elements, unit));
                }
                else { PipeGroup.Visibility = System.Windows.Visibility.Collapsed; }

                if (hasDuct)
                {
                    rbDucting.IsChecked = true;
                    SystemGrid.Visibility = System.Windows.Visibility.Collapsed;
                    DuctGroup.Visibility = System.Windows.Visibility.Visible;
                    DuctGrid.ItemsSource = QTOHelper.GetDuctInfoFromElements(doc, elements, unit);
                    allFittings.AddRange(QTOHelper.GetFittingInfoFromElements(doc, elements));
                    allInsulations.AddRange(QTOHelper.GetInsulationInfoFromElements(doc, elements, unit));
                }
                else { DuctGroup.Visibility = System.Windows.Visibility.Collapsed; }

                FittingGrid.ItemsSource = allFittings;

                if (rbPiping.IsChecked == true && pipingSystems != null)
                {
                    InsulationGrid.ItemsSource = allInsulations
                        .Where(i => pipingSystems.Any(p => p.SystemName == i.SystemName))
                        .OrderBy(i => i.SystemName).ThenBy(i => i.Diameter).ToList();
                }
                else if (rbDucting.IsChecked == true && ductingSystems != null)
                {
                    InsulationGrid.ItemsSource = allInsulations
                        .Where(i => ductingSystems.Any(d => d.SystemName == i.SystemName))
                        .OrderBy(i => i.SystemName).ThenBy(i => i.Diameter).ToList();
                }
                else
                {
                    InsulationGrid.ItemsSource = allInsulations;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading data from selection: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void PipeGrid_SelectionChanged(object sender, SelectionChangedEventArgs e) { }
        private void FittingGrid_SelectionChanged(object sender, SelectionChangedEventArgs e) { }
        private void Window_Loaded(object sender, RoutedEventArgs e) { }

        private void OnSelectByRectangleClick(object sender, RoutedEventArgs e)
        {
            this.Hide();
            UIDocument uidoc = _commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            string mode = rbPiping.IsChecked == true ? "Piping" : "Ducting";

            try
            {
                IList<Reference> pickedRefs = uidoc.Selection.PickObjects(
                    ObjectType.Element,
                    new QTOSelectionFilter(mode),
                    "Select objects using rectangle (Press Enter to finish)");

                List<Element> selectedElements = pickedRefs.Select(r => doc.GetElement(r)).ToList();

                this.Show();
                this.Topmost = true;
                _lastSelectedElements = pickedRefs.Select(r => doc.GetElement(r)).ToList();
                LoadFromElements(doc, selectedElements, CurrentLengthUnit, this, doc.ActiveView);
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                this.Show();
                this.Topmost = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error selecting objects: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                this.Show();
                this.Topmost = true;
            }
        }

        private enum SystemMode { None, Piping, Ducting }

        private void OnGridSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is DataGrid grid && grid.SelectedItem != null)
            {
                var prop = grid.SelectedItem.GetType().GetProperty("Reference");
                if (prop?.GetValue(grid.SelectedItem) is Reference reference)
                {
                    var uidoc = _commandData.Application.ActiveUIDocument;
                    uidoc.Selection.SetReferences(new List<Reference> { reference });
                    uidoc.Selection.SetElementIds(new List<ElementId> { reference.ElementId });
                }
            }
        }
    }

    public class SystemItem
    {
        public string SystemName { get; set; } = string.Empty;
        public bool IsSelected { get; set; }
    }
}