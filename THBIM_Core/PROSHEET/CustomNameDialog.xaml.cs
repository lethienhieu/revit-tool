using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace THBIM
{
    public partial class CustomNameDialog : Window
    {
        // Danh sách gốc bên trái
        private List<NameRuleItem> _allAvailableParams;

        // Danh sách hiển thị bên phải (Binding lên DataGrid)
        public ObservableCollection<NameRuleItem> SelectedRules { get; set; }

        // Kết quả trả về sau khi ấn OK (Ví dụ: "<Sheet Name>-<Sheet Number>")
        public string ResultFormatString { get; private set; }

        public CustomNameDialog(List<NameRuleItem> availableParams, ObservableCollection<NameRuleItem> existingRules = null)
        {
            InitializeComponent();
            _allAvailableParams = availableParams;

            // Nếu đã cấu hình trước đó thì load lại, nếu không thì tạo list trống
            SelectedRules = existingRules ?? new ObservableCollection<NameRuleItem>();

            // Đăng ký sự kiện thay đổi dữ liệu trong DataGrid để update Preview
            SelectedRules.CollectionChanged += (s, e) => UpdatePreview();
            foreach (var item in SelectedRules) item.PropertyChanged += (s, e) => UpdatePreview();

            DgNamingRules.ItemsSource = SelectedRules;

            RefreshAvailableList();
            UpdatePreview();
        }

        // Lọc danh sách Parameter bên trái dựa theo ô Search
        private void RefreshAvailableList()
        {
            string txt = TxtSearch.Text.Trim().ToLower();
            var filtered = _allAvailableParams.Where(p => p.ParameterName.ToLower().Contains(txt)).ToList();
            LstAvailable.ItemsSource = filtered;
        }

        private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e) => RefreshAvailableList();

        // THÊM Parameter (Bấm nút hoặc Double Click)
        private void AddSelectedParameter()
        {
            // Gom tất cả các dòng đang được bôi đen thành 1 danh sách
            var selectedItems = LstAvailable.SelectedItems.Cast<NameRuleItem>().ToList();

            if (selectedItems.Count == 0) return;

            foreach (var selected in selectedItems)
            {
                // Clone ra một object mới để chỉnh sửa độc lập
                var newItem = new NameRuleItem
                {
                    ParameterName = selected.ParameterName,
                    SampleValue = selected.SampleValue,
                    Separator = "-" // Gán mặc định dấu gạch ngang cho đẹp
                };

                // Bắt sự kiện gõ phím để update preview ngay lập tức
                newItem.PropertyChanged += (s, e) => UpdatePreview();

                SelectedRules.Add(newItem);
            }

            UpdatePreview();
        }

        private void BtnAdd_Click(object sender, RoutedEventArgs e) => AddSelectedParameter();
        private void LstAvailable_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e) => AddSelectedParameter();

        // XÓA Parameter
        private void BtnRemove_Click(object sender, RoutedEventArgs e)
        {
            if (DgNamingRules.SelectedItem is NameRuleItem selected)
            {
                SelectedRules.Remove(selected);
                UpdatePreview();
            }
        }

        // DI CHUYỂN LÊN XUỐNG
        private void BtnMoveUp_Click(object sender, RoutedEventArgs e)
        {
            int index = DgNamingRules.SelectedIndex;
            if (index > 0)
            {
                SelectedRules.Move(index, index - 1);
                UpdatePreview();
            }
        }

        private void BtnMoveDown_Click(object sender, RoutedEventArgs e)
        {
            int index = DgNamingRules.SelectedIndex;
            if (index >= 0 && index < SelectedRules.Count - 1)
            {
                SelectedRules.Move(index, index + 1);
                UpdatePreview();
            }
        }

        // HÀM QUAN TRỌNG: Nối chuỗi tạo Preview
        private void UpdatePreview()
        {
            if (TxtPreview == null) return;

            List<string> parts = new List<string>();
            List<string> formatParts = new List<string>(); // Để lưu lại chuỗi format nội bộ

            for (int i = 0; i < SelectedRules.Count; i++)
            {
                var rule = SelectedRules[i];
                // Chuỗi Preview hiển thị
                string partPreview = $"{rule.Prefix}{rule.SampleValue}{rule.Suffix}";
                // Chuỗi Format thực tế để tool xử lý (ví dụ: <Sheet Name>)
                string partFormat = $"{rule.Prefix}<{rule.ParameterName}>{rule.Suffix}";

                // Thêm Separator nếu KHÔNG phải là phần tử cuối cùng
                if (i < SelectedRules.Count - 1)
                {
                    partPreview += rule.Separator;
                    partFormat += rule.Separator;
                }

                parts.Add(partPreview);
                formatParts.Add(partFormat);
            }

            TxtPreview.Text = string.Join("", parts);
            ResultFormatString = string.Join("", formatParts);
        }

        // NÚT OK / CANCEL
        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
            this.Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }
    }
}