using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Xml.Serialization;
using DB = Autodesk.Revit.DB;
using ElementIdCompat = Autodesk.Revit.DB.ElementIdCompat;
using Win = System.Windows.Controls;

namespace THBIM
{
    public partial class THBIMSheetWindow : System.ComponentModel.INotifyPropertyChanged
    {

        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;
        private List<SheetItem> _masterList = new List<SheetItem>();
        public ObservableCollection<SheetItem> SheetItems { get; set; } = new ObservableCollection<SheetItem>();
        public ObservableCollection<SetItem> SetItems { get; set; } = new ObservableCollection<SetItem>();
        private ObservableCollection<NameRuleItem> _savedNamingRules = null;

        public ObservableCollection<string> ViewTypes { get; set; } = new ObservableCollection<string>();

        // --- THAY THẾ KHỐI NÀY ---
        private string _selectedViewType = "All View";
        public string SelectedViewType
        {
            get => _selectedViewType;
            set
            {
                _selectedViewType = value;
                // Khi code gán SelectedViewType = "All View", dòng lệnh này sẽ reng chuông 
                // báo cho giao diện WPF biết để tự động cập nhật hiển thị lên ComboBox.
                PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(SelectedViewType)));
            }
        }
        // ------------------------

        private async System.Threading.Tasks.Task InitSelectionTabAsync()
        {
            LoadingOverlay.Visibility = Visibility.Visible; // Hiện màn hình chờ

            InitFormatTab(); // Tab format load rất nhanh nên cứ để nguyên
            LoadSets();
            UpdateColumnVisibility();



            if (RbViews?.IsChecked == true)
                await LoadViewsAsync();
            else
                await LoadSheetsAsync();

            // Apply last profile SAU KHI sheets đã load xong
            ApplyLastProfile();

            LoadingOverlay.Visibility = Visibility.Collapsed; // Ẩn màn hình chờ
        }

        private void ApplyLastProfile()
        {
            string lastProfile = LoadLastProfileName();
            if (string.IsNullOrEmpty(lastProfile)) return;
            if (!_profilePaths.TryGetValue(lastProfile, out string path)) return;

            // Chỉ set tên combobox (suppress), rồi apply data thủ công
            // KHÔNG dùng SelectionChanged vì timing issue
            try
            {
                _suppressProfileSelection = true;
                foreach (System.Windows.Controls.ComboBoxItem ci in CboProfiles.Items)
                {
                    if (ci.Content?.ToString() == lastProfile)
                    {
                        CboProfiles.SelectedItem = ci;
                        break;
                    }
                }

                string content = File.ReadAllText(path);
                var data = ParseTxtProfile(content);
                if (data != null) ApplyTxtProfileToUI(data);
            }
            catch { }
            finally
            {
                _suppressProfileSelection = false;
            }
        }

        private string GetViewTypeName(DB.View v)
        {
            switch (v.ViewType)
            {
                case DB.ViewType.ThreeD: return "3D";
                case DB.ViewType.DraftingView: return "Drafting View";
                case DB.ViewType.Legend: return "Legend";
                case DB.ViewType.FloorPlan: return "Floor Plan";
                case DB.ViewType.CeilingPlan: return "Ceiling Plan";
                case DB.ViewType.EngineeringPlan: return "Engineering Plan";
                case DB.ViewType.Elevation: return "Elevation";
                case DB.ViewType.Section: return "Section";
                case DB.ViewType.Detail: return "Detail View";
                default: return v.ViewType.ToString();
            }
        }

        private void CboViewType_SelectionChanged(object sender, Win.SelectionChangedEventArgs e)
        {
            var combo = sender as Win.ComboBox;
            if (combo != null && combo.SelectedItem != null)
            {
                SelectedViewType = combo.SelectedItem.ToString();
                ApplyFilter();
            }
        }

        #region LOAD & REFRESH DATA
        private async System.Threading.Tasks.Task LoadSheetsAsync()
        {
            TxtLoadingStatus.Text = "Scanning sheets in project...";
            PbLoading.IsIndeterminate = true;
            await System.Threading.Tasks.Task.Delay(10);

            var sheets = new DB.FilteredElementCollector(_doc)
                .OfClass(typeof(DB.ViewSheet)).Cast<DB.ViewSheet>()
                .OrderBy(s => s.SheetNumber).ToList();

            // Pre-collect tất cả TitleBlocks 1 lần (thay vì mỗi sheet 1 collector)
            var titleBlockMap = new Dictionary<long, DB.Element>();
            foreach (var tb in new DB.FilteredElementCollector(_doc)
                .OfCategory(DB.BuiltInCategory.OST_TitleBlocks)
                .OfClass(typeof(DB.FamilyInstance)))
            {
                long ownerViewId = ElementIdCompat.GetValue(tb.OwnerViewId);
                if (!titleBlockMap.ContainsKey(ownerViewId))
                    titleBlockMap[ownerViewId] = tb;
            }

            _masterList = new List<SheetItem>();

            PbLoading.IsIndeterminate = false;
            PbLoading.Maximum = sheets.Count > 0 ? sheets.Count : 1;
            PbLoading.Value = 0;

            for (int i = 0; i < sheets.Count; i++)
            {
                var s = sheets[i];
                TxtLoadingStatus.Text = $"Loading Sheet/Views {i + 1} of {sheets.Count}...";
                PbLoading.Value = i + 1;

                // Lấy paper size từ pre-collected map
                titleBlockMap.TryGetValue(ElementIdCompat.GetValue(s.Id), out DB.Element cachedTb);
                string paperSize = cachedTb != null ? GetPaperSizeFromTitleBlock(cachedTb, s) : "None";

                _masterList.Add(new SheetItem
                {
                    Id = s.Id,
                    SheetNumber = s.SheetNumber ?? "",
                    SheetName = s.Name ?? "",
                    PaperSize = paperSize,
                    IsSelected = false
                });

                if (i % 50 == 0) await System.Threading.Tasks.Task.Delay(1);
            }

            ApplyFilter();
        }


        private string EvaluateFileNameFormat(DB.Element element, string formatString)
        {
            if (string.IsNullOrEmpty(formatString) || element == null) return "";

            string result = formatString;

            // Tìm tất cả các cụm nằm trong dấu < > (Ví dụ: <Sheet Name>)
            var matches = Regex.Matches(formatString, @"<(.*?)>");

            foreach (Match match in matches)
            {
                string paramName = match.Groups[1].Value;
                string paramValue = "";

                // Tìm Parameter tương ứng trong bản vẽ
                foreach (DB.Parameter p in element.Parameters)
                {
                    if (p.Definition != null && p.Definition.Name == paramName)
                    {
                        paramValue = p.AsValueString() ?? p.AsString() ?? "";
                        break;
                    }
                }

                // Ghi đè giá trị thật vào chỗ chứa công thức
                result = result.Replace(match.Value, paramValue);
            }

            // (Tùy chọn) Xóa các ký tự cấm trong tên file của Windows để tránh lỗi khi xuất file sau này
            char[] invalidChars = System.IO.Path.GetInvalidFileNameChars();
            foreach (char c in invalidChars)
            {
                result = result.Replace(c.ToString(), "_");
            }

            return result;
        }


        private async System.Threading.Tasks.Task LoadViewsAsync()
        {
            TxtLoadingStatus.Text = "Scanning views...";
            PbLoading.IsIndeterminate = true;
            await System.Threading.Tasks.Task.Delay(10);

            var views = new DB.FilteredElementCollector(_doc)
                .OfClass(typeof(DB.View)).Cast<DB.View>()
                .Where(v => !v.IsTemplate && !(v is DB.ViewSheet))
                .OrderBy(v => v.Name).ToList();

            _masterList = new List<SheetItem>();

            PbLoading.IsIndeterminate = false;
            PbLoading.Maximum = views.Count > 0 ? views.Count : 1;
            PbLoading.Value = 0;

            for (int i = 0; i < views.Count; i++)
            {
                var v = views[i];
                TxtLoadingStatus.Text = $"Loading view {i + 1} of {views.Count}...";
                PbLoading.Value = i + 1;

                // Trích xuất dữ liệu của View
                string discipline = "NA";
                var discParam = v.get_Parameter(DB.BuiltInParameter.VIEW_DISCIPLINE);
                if (discParam != null && discParam.HasValue) discipline = discParam.AsValueString();

                _masterList.Add(new SheetItem
                {
                    Id = v.Id,
                    SheetNumber = "",
                    SheetName = v.Name ?? "",
                    ViewType = GetViewTypeName(v),
                    ViewScale = v.Scale > 0 ? $"1 : {v.Scale}" : "Custom",
                    DetailLevel = v.DetailLevel.ToString(),
                    Discipline = discipline,
                    IsSelected = false
                });

                if (i % 10 == 0) await System.Threading.Tasks.Task.Delay(1);
            }

            // Ghi danh sách Type vào ComboBox
            var types = _masterList.Select(x => x.ViewType).Distinct().OrderBy(x => x).ToList();
            ViewTypes.Clear();
            ViewTypes.Add("All View");
            foreach (var t in types) ViewTypes.Add(t);
            SelectedViewType = "All View";

            ApplyFilter();
        }

        private void UpdateColumnVisibility()
        {
            // Kiểm tra an toàn xem RbSheets đã được khởi tạo chưa
            bool isSheet = RbSheets?.IsChecked == true;

            if (ColSheetNumber != null) ColSheetNumber.Visibility = isSheet ? Visibility.Visible : Visibility.Collapsed;
            if (ColRevision != null) ColRevision.Visibility = isSheet ? Visibility.Visible : Visibility.Collapsed;
            if (ColSize != null) ColSize.Visibility = isSheet ? Visibility.Visible : Visibility.Collapsed;

            if (ColViewType != null) ColViewType.Visibility = !isSheet ? Visibility.Visible : Visibility.Collapsed;
            if (ColViewScale != null) ColViewScale.Visibility = !isSheet ? Visibility.Visible : Visibility.Collapsed;
            if (ColDetailLevel != null) ColDetailLevel.Visibility = !isSheet ? Visibility.Visible : Visibility.Collapsed;
            if (ColDiscipline != null) ColDiscipline.Visibility = !isSheet ? Visibility.Visible : Visibility.Collapsed;

            if (ColName != null) ColName.Header = isSheet ? "Sheet Name" : "View Name";
        }
        private async void RbToggle_Checked(object sender, RoutedEventArgs e)
        {
            if (!this.IsLoaded || _doc == null) return;

            LoadingOverlay.Visibility = Visibility.Visible;

            // Gọi hàm xếp cột
            UpdateColumnVisibility();

            if (RbSheets.IsChecked == true)
                await LoadSheetsAsync();
            else if (RbViews.IsChecked == true)
                await LoadViewsAsync();

            LoadingOverlay.Visibility = Visibility.Collapsed;
        }


        private void LoadSets()
        {
            var sets = new DB.FilteredElementCollector(_doc)
                .OfClass(typeof(DB.ViewSheetSet)).Cast<DB.ViewSheetSet>()
                .Select(s => new SetItem { Id = s.Id, Name = s.Name })
                .OrderBy(s => s.Name).ToList();

            sets.Insert(0, new SetItem { Id = DB.ElementId.InvalidElementId, Name = "<Filter by Set>" });
            SetItems = new ObservableCollection<SetItem>(sets);
            CboSets.ItemsSource = SetItems;

            if (CboSets.SelectedIndex < 0) CboSets.SelectedIndex = 0;
            ApplyFilter();
        }

        private void RefreshSelectionUI()
        {
            string targetSetName = _handler.Req.SetName;
            if (_handler.Req.Type == RequestType.AddToSet)
            {
                var targetItem = SetItems.FirstOrDefault(x => x.Id == _handler.Req.SetId);
                if (targetItem != null) targetSetName = targetItem.Name;
            }
            else if (_handler.Req.Type == RequestType.DeleteSet) targetSetName = null;

            LoadSets();

            if (!string.IsNullOrEmpty(targetSetName))
            {
                var match = SetItems.FirstOrDefault(x => x.Name == targetSetName);
                if (match != null) CboSets.SelectedItem = match;
            }
            else CboSets.SelectedIndex = 0;

            foreach (var item in _masterList) item.IsSelected = false; // Reset check
            ApplyFilter();
        }
        #endregion



        private void HeaderCustomNameBtn_Click(object sender, RoutedEventArgs e)
        {
            // Vẫn lấy 1 bản vẽ mẫu (ưu tiên bản vẽ đang chọn, nếu không thì lấy cái đầu tiên) để trích xuất Parameter
            var sampleSheet = _masterList.FirstOrDefault(x => x.IsSelected) ?? _masterList.FirstOrDefault();

            if (sampleSheet == null)
            {
                MessageBox.Show("The list is empty!", "Notice", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var sampleElement = _doc.GetElement(sampleSheet.Id);
            if (sampleElement == null) return;

            // Trích xuất Parameter của bản vẽ mẫu
            List<NameRuleItem> availableParams = new List<NameRuleItem>();
            foreach (DB.Parameter p in sampleElement.Parameters)
            {
                if (p.Definition != null && !string.IsNullOrWhiteSpace(p.Definition.Name))
                {
                    string val = p.AsValueString() ?? p.AsString() ?? "";
                    if (string.IsNullOrWhiteSpace(val)) val = "[Empty]";

                    availableParams.Add(new NameRuleItem
                    {
                        ParameterName = p.Definition.Name,
                        SampleValue = val
                    });
                }
            }

            // Lọc trùng và sắp xếp A-Z
            availableParams = availableParams
                .GroupBy(x => x.ParameterName)
                .Select(g => g.First())
                .OrderBy(x => x.ParameterName)
                .ToList();

            // Mở bảng cài đặt (Dialog)
            var dialog = new CustomNameDialog(availableParams, _savedNamingRules);

            if (dialog.ShowDialog() == true)
            {
                // LƯU LẠI CẤU HÌNH NGAY LẬP TỨC ĐỂ LẦN SAU MỞ LÊN VẪN CÒN
                _savedNamingRules = dialog.SelectedRules;

                string rawFormat = dialog.ResultFormatString; // Chuỗi thô: <Sheet Number>-<Sheet Name>

                // ÁP DỤNG HÀNG LOẠT CHO TOÀN BỘ _masterList
                foreach (var item in _masterList)
                {
                    var elem = _doc.GetElement(item.Id);
                    if (elem != null)
                    {
                        // Gọi hàm dịch để lấy tên thật
                        item.CustomFileName = EvaluateFileNameFormat(elem, rawFormat);
                    }
                }

                // Cập nhật lại giao diện DataGrid
                if (DgSheets != null)
                {
                    DgSheets.Items.Refresh();
                }
            }
        }


        private string GetPaperSizeFromTitleBlock(DB.Element titleBlock, DB.ViewSheet sheet)
        {
            if (titleBlock == null) return "None";

            DB.BoundingBoxXYZ bbox = titleBlock.get_BoundingBox(sheet);
            if (bbox == null) return "Unknown";

            double widthMm = Math.Round((bbox.Max.X - bbox.Min.X) * 304.8);
            double heightMm = Math.Round((bbox.Max.Y - bbox.Min.Y) * 304.8);

            double longEdge = Math.Max(widthMm, heightMm);
            double shortEdge = Math.Min(widthMm, heightMm);

            double tolerance = 5.0;
            if (Math.Abs(longEdge - 1189) <= tolerance && Math.Abs(shortEdge - 841) <= tolerance) return "A0";
            if (Math.Abs(longEdge - 841) <= tolerance && Math.Abs(shortEdge - 594) <= tolerance) return "A1";
            if (Math.Abs(longEdge - 594) <= tolerance && Math.Abs(shortEdge - 420) <= tolerance) return "A2";
            if (Math.Abs(longEdge - 420) <= tolerance && Math.Abs(shortEdge - 297) <= tolerance) return "A3";
            if (Math.Abs(longEdge - 297) <= tolerance && Math.Abs(shortEdge - 210) <= tolerance) return "A4";

            return $"{longEdge}x{shortEdge}";
        }

        #region FILTER & CHECKBOX
        private void ApplyFilter()
        {
            IEnumerable<SheetItem> query = _masterList;

            var selectedSet = CboSets.SelectedItem as SetItem;
            if (selectedSet != null && selectedSet.Id != DB.ElementId.InvalidElementId)
            {
                var vss = _doc.GetElement(selectedSet.Id) as DB.ViewSheetSet;
                if (vss != null)
                {
                    var viewIds = vss.Views.Cast<DB.View>().Select(v => v.Id).ToHashSet();
                    query = query.Where(x => viewIds.Contains(x.Id));
                }
            }

            string txt = TxtSearch.Text.Trim().ToLower();
            if (!string.IsNullOrEmpty(txt))
                query = query.Where(x => x.SheetNumber.ToLower().Contains(txt) || x.SheetName.ToLower().Contains(txt));

            // THÊM ĐOẠN NÀY ĐỂ LỌC THEO COMBOBOX
            if (RbViews?.IsChecked == true && SelectedViewType != "All View" && !string.IsNullOrEmpty(SelectedViewType))
            {
                query = query.Where(x => x.ViewType == SelectedViewType);
            }

            SheetItems.Clear();
            foreach (var item in query) SheetItems.Add(item);
            if (DgSheets != null) DgSheets.Items.Refresh();
            UpdateCount();
        }

        private void TxtSearch_TextChanged(object sender, Win.TextChangedEventArgs e) => ApplyFilter();
        private void CboSets_SelectionChanged(object sender, Win.SelectionChangedEventArgs e) => ApplyFilter();

        private void HeaderCheckBox_Click(object sender, RoutedEventArgs e)
        {
            bool val = (sender as Win.CheckBox).IsChecked ?? false;
            foreach (var i in SheetItems) i.IsSelected = val;
            UpdateCount();
        }

        private void RowCheckBox_Click(object sender, RoutedEventArgs e)
        {
            bool val = (sender as Win.CheckBox).IsChecked ?? false;
            var item = (sender as Win.CheckBox).DataContext as SheetItem;
            if (DgSheets.SelectedItems.Count > 1 && DgSheets.SelectedItems.Contains(item))
                foreach (SheetItem i in DgSheets.SelectedItems) i.IsSelected = val;
            UpdateCount();
        }

        private void UpdateCount() { if (TxtCount != null) TxtCount.Text = $"{_masterList.Count(x => x.IsSelected)} selected"; }
        #endregion

        #region SET MANAGEMENT (UI)
        private void CboSaveSet_SelectionChanged(object sender, Win.SelectionChangedEventArgs e)
        {
            // Bỏ qua nếu ComboBox chưa load hoặc đang chọn lại tiêu đề
            if (CboSaveSet == null || CboSaveSet.SelectedIndex <= 0) return;

            int selectedIndex = CboSaveSet.SelectedIndex;

            // 1. Trả về Index 0 ngay lập tức để ComboBox luôn hiển thị chữ "Save V/S Set" trên bề mặt
            CboSaveSet.SelectedIndex = 0;

            // 2. Tái sử dụng lại 100% các hàm chức năng cũ của bạn
            if (selectedIndex == 1)
            {
                MenuNewSet_Click(null, null);
            }
            else if (selectedIndex == 2)
            {
                MenuAddToExisting_Click(null, null);
            }
            else if (selectedIndex == 3)
            {
                MenuDeleteSet_Click(null, null);
            }
        }

        private void MenuNewSet_Click(object sender, RoutedEventArgs e)
        {
            var ids = _masterList.Where(x => x.IsSelected).Select(x => x.Id).ToList();
            if (!ids.Any()) { MessageBox.Show("No sheets selected!"); return; }

            string name = InputDialog("Create New Set", "Enter Set name:");
            if (string.IsNullOrWhiteSpace(name)) return;

            SendRequest(RequestType.CreateSet, ids, name);
        }

        private void MenuAddToExisting_Click(object sender, RoutedEventArgs e)
        {
            var ids = _masterList.Where(x => x.IsSelected).Select(x => x.Id).ToList();
            if (!ids.Any()) { MessageBox.Show("No sheets selected!"); return; }

            var existingSets = SetItems.Where(x => x.Id != DB.ElementId.InvalidElementId).ToList();
            var target = SelectSetDialog(existingSets);
            if (target == null) return;

            SendRequest(RequestType.AddToSet, ids, target.Name, target.Id);
        }

        private void MenuDeleteSet_Click(object sender, RoutedEventArgs e)
        {
            var set = CboSets.SelectedItem as SetItem;
            if (set == null || set.Id == DB.ElementId.InvalidElementId) return;
            if (MessageBox.Show($"Delete Set '{set.Name}'?", "Confirm", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                SendRequest(RequestType.DeleteSet, null, null, set.Id);
        }

        private void SendRequest(RequestType type, List<DB.ElementId> ids, string name = null, DB.ElementId setId = null)
        {
            _handler.Req.Type = type;
            _handler.Req.Ids = ids;
            _handler.Req.SetName = name;
            _handler.Req.SetId = setId;
            _exEvent.Raise();
        }

        // Popup Helpers
        private string InputDialog(string title, string msg)
        {
            var w = new Window { Title = title, Width = 300, Height = 140, WindowStartupLocation = WindowStartupLocation.CenterScreen, ResizeMode = ResizeMode.NoResize };
            var p = new Win.StackPanel { Margin = new Thickness(10) };
            var t = new Win.TextBox { Margin = new Thickness(0, 5, 0, 10) };
            var b = new Win.Button { Content = "OK", IsDefault = true, HorizontalAlignment = HorizontalAlignment.Right, Width = 60 };
            b.Click += (s, e) => w.DialogResult = true;
            p.Children.Add(new Win.TextBlock { Text = msg }); p.Children.Add(t); p.Children.Add(b); w.Content = p;
            return w.ShowDialog() == true ? t.Text : null;
        }

        private SetItem SelectSetDialog(List<SetItem> items)
        {
            var w = new Window { Title = "Select Set", Width = 300, Height = 140, WindowStartupLocation = WindowStartupLocation.CenterScreen, ResizeMode = ResizeMode.NoResize };
            var p = new Win.StackPanel { Margin = new Thickness(10) };
            var c = new Win.ComboBox { ItemsSource = items, DisplayMemberPath = "Name", SelectedIndex = 0, Margin = new Thickness(0, 5, 0, 10) };
            var b = new Win.Button { Content = "OK", IsDefault = true, HorizontalAlignment = HorizontalAlignment.Right, Width = 60 };
            b.Click += (s, e) => w.DialogResult = true;
            p.Children.Add(new Win.TextBlock { Text = "Add to:" }); p.Children.Add(c); p.Children.Add(b); w.Content = p;
            return w.ShowDialog() == true ? c.SelectedItem as SetItem : null;
        }
        #endregion
    }
}