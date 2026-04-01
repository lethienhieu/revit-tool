using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace THBIM
{
    public partial class OverlapWindow : Window
    {
        private UIDocument _uidoc;
        private ExternalEvent _exEvent;
        private OverlapHandler _handler;
        private bool _isInitializing = true;

        public OverlapWindow(UIDocument uidoc, ExternalEvent exEvent, OverlapHandler handler)
        {
            InitializeComponent();
            _uidoc = uidoc;
            _exEvent = exEvent;
            _handler = handler;
            _handler.MyWindow = this;

            LoadCategories();
            _isInitializing = false;
            UpdateUIFromCache();
        }

        private void LoadCategories()
        {
            List<CategoryOption> cats = CheckOverlapLogic.GetAllModelCategories(_uidoc.Document);
            cmbCategories.ItemsSource = cats;

            var defaultCat = cats.FirstOrDefault(x => x.BIC == BuiltInCategory.OST_StructuralFoundation);
            if (defaultCat != null)
                cmbCategories.SelectedItem = defaultCat;
            else if (cats.Count > 0)
                cmbCategories.SelectedIndex = 0;

            if (cats.Count == 0)
            {
                btnScan.IsEnabled = false;
                MessageBox.Show("No valid 3D model categories found in view.", "Info");
            }
        }

        private void CmbCategories_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isInitializing) return;
            UpdateUIFromCache();
        }

        private void UpdateUIFromCache()
        {
            CategoryOption selectedOpt = cmbCategories.SelectedItem as CategoryOption;
            if (selectedOpt == null) return;

            if (CallUIOverlap.GlobalCache.ContainsKey(selectedOpt.BIC))
            {
                var cachedList = CallUIOverlap.GlobalCache[selectedOpt.BIC];
                List<OverlapGroup> validGroups = new List<OverlapGroup>();
                bool hasChanges = false;

                foreach (var group in cachedList)
                {
                    bool isValidGroup = true;
                    foreach (var item in group.Items)
                    {
                        if (_uidoc.Document.GetElement(item.Id) == null)
                        {
                            isValidGroup = false;
                            break;
                        }
                    }

                    if (isValidGroup) validGroups.Add(group);
                    else hasChanges = true;
                }

                if (hasChanges) CallUIOverlap.GlobalCache[selectedOpt.BIC] = validGroups;

                tvResults.ItemsSource = null;
                tvResults.ItemsSource = validGroups;
            }
            else
            {
                tvResults.ItemsSource = null;
            }
        }

        // --- SỰ KIỆN CLICK VÀO LIST: CHỈ HIGHLIGHT (KHÔNG ZOOM) ---
        private void TvResults_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (_uidoc == null) return;

            var selectedObj = tvResults.SelectedItem;
            List<ElementId> idsToSelect = new List<ElementId>();

            try
            {
                if (selectedObj is OverlapItem item)
                {
                    idsToSelect.Add(item.Id);
                }
                else if (selectedObj is OverlapGroup group)
                {
                    foreach (var subItem in group.Items) idsToSelect.Add(subItem.Id);
                }

                if (idsToSelect.Count > 0)
                {
                    // CHỈ GỌI LỆNH SELECT (Sáng xanh)
                    _uidoc.Selection.SetElementIds(idsToSelect);

                    // ĐÃ BỎ LỆNH ShowElements (Zoom) TẠI ĐÂY
                }
            }
            catch { }
        }

        private void BtnScan_Click(object sender, RoutedEventArgs e)
        {
            CategoryOption selectedOpt = cmbCategories.SelectedItem as CategoryOption;
            if (selectedOpt == null)
            {
                MessageBox.Show("Please select a category first.");
                return;
            }
            _handler.SelectedCategory = selectedOpt.BIC;
            _exEvent.Raise();
        }

        // --- NÚT SHOW: LÚC NÀY MỚI ZOOM ---
        private void BtnShow_Click(object sender, RoutedEventArgs e)
        {
            var selectedObj = tvResults.SelectedItem;
            List<ElementId> idsToZoom = new List<ElementId>();

            if (selectedObj is OverlapItem item)
            {
                idsToZoom.Add(item.Id);
            }
            else if (selectedObj is OverlapGroup group)
            {
                foreach (var subItem in group.Items) idsToZoom.Add(subItem.Id);
            }

            if (idsToZoom.Count > 0)
            {
                try
                {
                    _uidoc.Selection.SetElementIds(idsToZoom); // Select lại cho chắc
                    _uidoc.ShowElements(idsToZoom); // ZOOM TỚI ĐỐI TƯỢNG
                }
                catch { }
            }
            else
            {
                MessageBox.Show("Please select an item to show.");
            }
        }

        public void UpdateResults(List<OverlapGroup> groups)
        {
            CategoryOption selectedOpt = cmbCategories.SelectedItem as CategoryOption;
            if (selectedOpt != null)
            {
                if (CallUIOverlap.GlobalCache.ContainsKey(selectedOpt.BIC))
                    CallUIOverlap.GlobalCache[selectedOpt.BIC] = groups;
                else
                    CallUIOverlap.GlobalCache.Add(selectedOpt.BIC, groups);
            }

            if (groups.Count > 0)
            {
                tvResults.ItemsSource = groups;
                MessageBox.Show($"Found {groups.Count} conflicts.\nSaved to cache.", "Done");
            }
            else
            {
                tvResults.ItemsSource = null;
                MessageBox.Show("No significant overlaps found!", "Clean");
            }
        }
    }
}