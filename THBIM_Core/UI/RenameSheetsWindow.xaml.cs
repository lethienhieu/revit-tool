using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Text.RegularExpressions;
using System.CodeDom.Compiler;
#nullable disable

namespace THBIM
{
    public partial class RenameSheetsWindow : Window
    {

        private readonly List<ViewSheet> _sheets;
        private readonly ObservableCollection<RenameSheetItem> _displaySheets;
        private ObservableCollection<string> _sheetParameters;

        public RenameSheetsWindow(List<ViewSheet> sheets)
        {
            InitializeComponent();
            _sheets = sheets ?? new List<ViewSheet>();

            _displaySheets = new ObservableCollection<RenameSheetItem>(
                _sheets.Select((s, idx) => new RenameSheetItem
                {
                    SheetId = s.Id.GetValue(),
                    OriginalNumber = s.SheetNumber,
                    DisplayNumber = s.SheetNumber,
                    OriginalName = s.Name,
                    DisplayName = s.Name,
                    NumberingValue = "",
                    OrderKey = idx
                })
            );
            dgSheets.ItemsSource = _displaySheets;
            ApplyOrderKeySortOnly();

            // build parameter list
            _sheetParameters = new ObservableCollection<string>();
            if (_sheets.Count > 0)
            {
                var first = _sheets.First();
                foreach (Parameter p in first.Parameters)
                {
                    if (p?.Definition == null) continue;
                    var name = p.Definition.Name;
                    if (!_sheetParameters.Contains(name)) _sheetParameters.Add(name);
                }
            }
            cbNumberingParam.ItemsSource = _sheetParameters;
            cbNumberingParam.SelectedIndex = -1; // để trống khi mở
        }

        public string FindText => txtFind.Text;
        public string ReplaceText => txtReplace.Text;
        public bool IsRenameNumber => rbShowNumber.IsChecked == true;

        private void FindReplace_TextChanged(object sender, TextChangedEventArgs e)
        {
            string find = txtFind.Text ?? "";
            string replace = txtReplace.Text ?? "";
            foreach (var item in _displaySheets)
            {
                if (!string.IsNullOrEmpty(find))
                {
                    item.DisplayNumber = (item.OriginalNumber ?? "").Replace(find, replace);
                    item.DisplayName = (item.OriginalName ?? "").Replace(find, replace);
                }
                else
                {
                    item.DisplayNumber = item.OriginalNumber;
                    item.DisplayName = item.OriginalName;
                }

            }
        }

        private void rbShowNumber_Checked(object sender, RoutedEventArgs e)
        {
            if (colSheetNumber == null || colSheetName == null) return;
            colSheetNumber.Visibility = System.Windows.Visibility.Visible;
            colSheetName.Visibility = System.Windows.Visibility.Collapsed;
        }

        private void rbShowName_Checked(object sender, RoutedEventArgs e)
        {
            if (colSheetNumber == null || colSheetName == null) return;
            colSheetNumber.Visibility = System.Windows.Visibility.Collapsed;
            colSheetName.Visibility = System.Windows.Visibility.Visible;
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            bool ok = ApplyChanges(showMessage: false);
            if (ok)
            {
                DialogResult = true;
                Close();
            }
        }

        private void Apply_Click(object sender, RoutedEventArgs e)
        {
            ApplyChanges(showMessage: true);
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private bool ApplyChanges(bool showMessage)
        {
            var doc = _sheets.FirstOrDefault()?.Document;
            if (doc == null) return false;

            try
            {
                using (Transaction t = new Transaction(doc, "Rename Sheets + Numbering"))
                {
                    t.Start();

                    if (IsRenameNumber)
                    {
                        // === ĐỔI SHEET NUMBER THEO THỨ TỰ AN TOÀN TRONG NHÓM ===
                        // Gom cặp (Sheet, OldNum, NewNum) chỉ cho các sheet thực sự đổi số
                        var pairs = _displaySheets
                            .Select(it =>
                            {
                                var sh = doc.GetElement(ElementIdCompat.CreateId(it.SheetId)) as ViewSheet;
                                return new { Sheet = sh, OldNum = sh?.SheetNumber ?? "", NewNum = it.DisplayNumber ?? "" };
                            })
                            .Where(p => p.Sheet != null && p.OldNum != p.NewNum)
                            .ToList();

                        // Hàm lấy số đuôi
                        Func<string, int?> tail = s =>
                        {
                            var m = Regex.Match(s ?? "", @"(\d+)$");
                            return m.Success ? int.Parse(m.Groups[1].Value) : (int?)null;
                        };

                        bool allParsed = pairs.All(p => tail(p.OldNum).HasValue && tail(p.NewNum).HasValue);
                        if (allParsed)
                        {
                            // Nếu toàn bộ là tăng (01->02, 10->11) thì đổi từ lớn -> nhỏ
                            bool increasing = pairs.All(p => tail(p.NewNum).Value > tail(p.OldNum).Value);
                            // Nếu toàn bộ là giảm (10->09, 02->01) thì đổi từ nhỏ -> lớn
                            bool decreasing = pairs.All(p => tail(p.NewNum).Value < tail(p.OldNum).Value);

                            if (increasing)
                                pairs = pairs.OrderByDescending(p => tail(p.OldNum).Value).ToList();
                            else if (decreasing)
                                pairs = pairs.OrderBy(p => tail(p.OldNum).Value).ToList();
                            // Nếu lẫn lộn thì giữ nguyên thứ tự hiện tại (OrderKey đã chi phối UI)
                        }

                        // Đặt số mới theo thứ tự đã sắp
                        foreach (var p in pairs)
                            p.Sheet.SheetNumber = p.NewNum;
                    }
                    else
                    {
                        // Đổi NAME: không yêu cầu duy nhất -> set trực tiếp
                        foreach (var item in _displaySheets)
                        {
                            var sheet = doc.GetElement(ElementIdCompat.CreateId(item.SheetId)) as ViewSheet;
                            if (sheet != null) sheet.Name = item.DisplayName;
                        }
                    }

                    // Set Numbering parameter nếu có chọn và có giá trị
                    if (cbNumberingParam.SelectedItem != null)
                    {
                        string paramName = cbNumberingParam.SelectedItem.ToString();
                        foreach (var item in _displaySheets)
                        {
                            if (string.IsNullOrEmpty(item.NumberingValue)) continue;
                            var sheet = doc.GetElement(ElementIdCompat.CreateId(item.SheetId)) as ViewSheet;
                            var p = sheet?.LookupParameter(paramName);
                            if (p != null && !p.IsReadOnly)
                                p.Set(item.NumberingValue);
                        }
                    }

                    t.Commit();
                }

                // Chốt lại Original = Display (để lần sau preview đúng)
                foreach (var it in _displaySheets)
                {
                    it.OriginalNumber = it.DisplayNumber;
                    it.OriginalName = it.DisplayName;
                }

                if (showMessage)
                    MessageBox.Show("Sheets have been updated in project.", "Apply",
                                    MessageBoxButton.OK, MessageBoxImage.Information);

                return true;
            }
            catch
            {
                return false;
            }
        }



        private void btnNumbering_Click(object sender, RoutedEventArgs e)
        {
            // Đọc Start & Digits
            int start = 1, digits = 2;
            int.TryParse(txtNumStart?.Text, out start);
            int.TryParse(txtNumDigits?.Text, out digits);
            if (digits < 1) digits = 1;              // tối thiểu 1 chữ số
            if (start < 0) start = 0;                // cho phép 0 nếu muốn

            // Đánh số theo thứ tự hiển thị (OrderKey)
            var items = _displaySheets.OrderBy(x => x.OrderKey).ToList();

            for (int i = 0; i < items.Count; i++)
            {
                string numbering = (start + i).ToString("D" + digits);
                items[i].NumberingValue = numbering;
                items[i].PreviewParameterName = cbNumberingParam.SelectedItem?.ToString();
                items[i].PreviewParameterValue = numbering;
            }

            dgSheets.Items.Refresh();
        }


        private void cbNumberingParam_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cbNumberingParam.SelectedItem == null) return;
            string pname = cbNumberingParam.SelectedItem.ToString();
            if (string.IsNullOrWhiteSpace(pname)) return;

            var doc = _sheets.FirstOrDefault()?.Document;
            if (doc == null) return;

            foreach (var item in _displaySheets)
            {
                var sheet = doc.GetElement(ElementIdCompat.CreateId(item.SheetId)) as ViewSheet;
                string v = ReadRevitParam(sheet?.LookupParameter(pname));
                item.NumberingValue = v;
                item.PreviewParameterName = pname;
                item.PreviewParameterValue = v;
            }
            dgSheets.Items.Refresh();
        }

        // ======= 4 nút điều hướng tác động lên các hàng đang chọn =======
        private void MoveTop_Click(object sender, RoutedEventArgs e)
        {
            var sel = dgSheets.SelectedItems.Cast<RenameSheetItem>().ToList();
            if (sel.Count == 0) return;
            var list = _displaySheets.ToList();

            // remove & insert lên đầu, giữ nguyên thứ tự các item đã chọn
            list.RemoveAll(i => sel.Contains(i));
            list.InsertRange(0, sel);
            Rebuild(list, sel);
        }

        private void MoveBottom_Click(object sender, RoutedEventArgs e)
        {
            var sel = dgSheets.SelectedItems.Cast<RenameSheetItem>().ToList();
            if (sel.Count == 0) return;
            var list = _displaySheets.ToList();

            list.RemoveAll(i => sel.Contains(i));
            list.AddRange(sel);
            Rebuild(list, sel);
        }

        private void MoveUp_Click(object sender, RoutedEventArgs e)
        {
            var sel = dgSheets.SelectedItems.Cast<RenameSheetItem>().ToHashSet();
            if (sel.Count == 0) return;

            var list = _displaySheets.ToList();
            for (int i = 1; i < list.Count; i++)
            {
                if (sel.Contains(list[i]) && !sel.Contains(list[i - 1]))
                {
                    (list[i - 1], list[i]) = (list[i], list[i - 1]);
                }
            }
            Rebuild(list, sel.ToList());
        }

        private void MoveDown_Click(object sender, RoutedEventArgs e)
        {
            var sel = dgSheets.SelectedItems.Cast<RenameSheetItem>().ToHashSet();
            if (sel.Count == 0) return;

            var list = _displaySheets.ToList();
            for (int i = list.Count - 2; i >= 0; i--)
            {
                if (sel.Contains(list[i]) && !sel.Contains(list[i + 1]))
                {
                    (list[i + 1], list[i]) = (list[i], list[i + 1]);
                }
            }
            Rebuild(list, sel.ToList());
        }

        private void Rebuild(List<RenameSheetItem> inOrder, IList<RenameSheetItem> toSelect)
        {
            // cập nhật OrderKey = 0..n-1 (STT không đổi)
            for (int i = 0; i < inOrder.Count; i++) inOrder[i].OrderKey = i;

            _displaySheets.Clear();
            foreach (var it in inOrder) _displaySheets.Add(it);

            ApplyOrderKeySortOnly();

            // giữ selection và cuộn tới item cuối cùng trong nhóm
            dgSheets.UnselectAll();
            foreach (var it in toSelect) dgSheets.SelectedItems.Add(it);
            if (toSelect.Count > 0)
            {
                var last = toSelect.Last();
                dgSheets.CurrentItem = last;
                dgSheets.ScrollIntoView(last);
            }
        }

        private void ApplyOrderKeySortOnly()
        {
            var view = (ListCollectionView)CollectionViewSource.GetDefaultView(dgSheets.ItemsSource);
            view.CustomSort = null;
            view.SortDescriptions.Clear();
            view.SortDescriptions.Add(new SortDescription(nameof(RenameSheetItem.OrderKey), ListSortDirection.Ascending));
            view.Refresh();
        }

        // ======= Sort header (cập nhật OrderKey, không xung đột) =======
        private void dgSheets_Sorting(object sender, DataGridSortingEventArgs e)
        {
            e.Handled = true;

            var newDir = (e.Column.SortDirection != ListSortDirection.Ascending)
                         ? ListSortDirection.Ascending
                         : ListSortDirection.Descending;
            e.Column.SortDirection = newDir;

            string prop = (e.Column as DataGridTemplateColumn)?.SortMemberPath
                       ?? (e.Column as DataGridTextColumn)?.SortMemberPath;
            if (string.IsNullOrEmpty(prop)) return;

            var icomp = StringComparer.CurrentCultureIgnoreCase;
            IEnumerable<RenameSheetItem> ordered = _displaySheets;

            if (prop == nameof(RenameSheetItem.DisplayNumber))
                ordered = (newDir == ListSortDirection.Ascending)
                    ? _displaySheets.OrderBy(x => x.DisplayNumber, icomp)
                    : _displaySheets.OrderByDescending(x => x.DisplayNumber, icomp);
            else if (prop == nameof(RenameSheetItem.DisplayName))
                ordered = (newDir == ListSortDirection.Ascending)
                    ? _displaySheets.OrderBy(x => x.DisplayName, icomp)
                    : _displaySheets.OrderByDescending(x => x.DisplayName, icomp);
            else if (prop == nameof(RenameSheetItem.NumberingValue))
                ordered = (newDir == ListSortDirection.Ascending)
                    ? _displaySheets.OrderBy(x => x.NumberingValue, icomp)
                    : _displaySheets.OrderByDescending(x => x.NumberingValue, icomp);

            var list = ordered.ToList();
            for (int i = 0; i < list.Count; i++) list[i].OrderKey = i;

            _displaySheets.Clear();
            foreach (var it in list) _displaySheets.Add(it);

            ApplyOrderKeySortOnly();
        }

        // ======= helpers =======
        private static string ReadRevitParam(Autodesk.Revit.DB.Parameter p)
        {
            if (p == null) return "";
            try
            {
                string vs = p.AsValueString();
                if (!string.IsNullOrEmpty(vs)) return vs;

                switch (p.StorageType)
                {
                    case StorageType.String: return p.AsString() ?? "";
                    case StorageType.Integer: return p.AsInteger().ToString(CultureInfo.InvariantCulture);
                    case StorageType.Double: return p.AsDouble().ToString(CultureInfo.InvariantCulture);
                    case StorageType.ElementId: return p.AsElementId().GetValue().ToString(CultureInfo.InvariantCulture);
                    default: return "";
                }
            }
            catch { return ""; }
        }
        private void btnDetectPrefix_Click(object sender, RoutedEventArgs e)
        {
            // Lấy mẫu từ hàng đang chọn, nếu không có thì lấy hàng đầu tiên
            string sample = "";
            if (dgSheets.SelectedItems.Count > 0)
                sample = (dgSheets.SelectedItems[0] as RenameSheetItem)?.DisplayNumber ?? "";
            else
                sample = _displaySheets.FirstOrDefault()?.DisplayNumber ?? "";

            txtSNPrefix.Text = ExtractSheetPrefix(sample);
        }

        private void btnApplySeq_Click(object sender, RoutedEventArgs e)
        {
            // Prefix: nếu để trống, tự lấy từ mẫu đầu tiên
            string prefix = txtSNPrefix.Text?.Trim();
            if (string.IsNullOrEmpty(prefix))
            {
                var sample = _displaySheets.FirstOrDefault()?.DisplayNumber ?? "";
                prefix = ExtractSheetPrefix(sample);
                txtSNPrefix.Text = prefix;
            }

            // Start & Digits
            int start = 1;
            int digits = 2;
            int.TryParse(txtSNStart.Text, out start);
            int.TryParse(txtSNDigits.Text, out digits);
            if (digits < 1) digits = 1;

            // Tập áp dụng: selected hay toàn bộ – theo thứ tự đang hiển thị (OrderKey)
            var target = (cbSNSelected.IsChecked == true
                ? dgSheets.SelectedItems.Cast<RenameSheetItem>()
                : _displaySheets).OrderBy(x => x.OrderKey).ToList();

            for (int i = 0; i < target.Count; i++)
            {
                string suffix = (start + i).ToString("D" + digits);
                target[i].DisplayNumber = prefix + suffix;
            }

            // Bật radio Number để khi Apply thì ghi vào SheetNumber
            rbShowNumber.IsChecked = true;

            dgSheets.Items.Refresh();
        }
        private static string ExtractSheetPrefix(string number)
        {
            if (string.IsNullOrEmpty(number)) return "";
            // Bóc dãy số ở cuối: "....-L2-01" -> prefix = "....-L2-"
            var m = Regex.Match(number, @"^(.*?)(\d+)$");
            return m.Success ? m.Groups[1].Value : number;
        }

    }

    public class RenameSheetItem : INotifyPropertyChanged
    {
        public long SheetId { get; set; }

        public string OriginalNumber { get; set; }
        public string OriginalName { get; set; }

        private string _displayNumber;
        public string DisplayNumber
        {
            get => _displayNumber;
            set { _displayNumber = value; OnPropertyChanged(nameof(DisplayNumber)); }
        }

        private string _displayName;
        public string DisplayName
        {
            get => _displayName;
            set { _displayName = value; OnPropertyChanged(nameof(DisplayName)); }
        }

        private string _numberingValue;
        public string NumberingValue
        {
            get => _numberingValue;
            set { _numberingValue = value; OnPropertyChanged(nameof(NumberingValue)); }
        }

        private int _orderKey;
        public int OrderKey
        {
            get => _orderKey;
            set { _orderKey = value; OnPropertyChanged(nameof(OrderKey)); }
        }

        private string _previewParameterName;
        public string PreviewParameterName
        {
            get => _previewParameterName;
            set { _previewParameterName = value; OnPropertyChanged(nameof(PreviewParameterName)); }
        }

        private string _previewParameterValue;
        public string PreviewParameterValue
        {
            get => _previewParameterValue;
            set { _previewParameterValue = value; OnPropertyChanged(nameof(PreviewParameterValue)); }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string n) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));
    }
}
