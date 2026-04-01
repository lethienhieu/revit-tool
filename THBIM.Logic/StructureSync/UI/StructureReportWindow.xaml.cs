using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace THBIM.Views
{
    public partial class StructureReportWindow : Window
    {
        private ExternalEvent _exEvent;
        private StructureSyncRequestHandler _handler;
        private List<RelationshipItem> _sourceRelationships; // Tham chiếu đến list gốc để xóa

        // [QUAN TRỌNG] Biến cờ hiệu để chặn vòng lặp Selection
        private bool _isProgrammaticSelection = false;

        public StructureReportWindow(List<ReportItem> reports, ExternalEvent exEvent, StructureSyncRequestHandler handler, List<RelationshipItem> sourceRelationships)
        {
            InitializeComponent();
            _exEvent = exEvent;
            _handler = handler;
            _sourceRelationships = sourceRelationships;

            DgReport.ItemsSource = reports;

            // Rule: if multiple sets are ticked, hide the Add button.
            int tickedCount = _sourceRelationships == null ? 0 : _sourceRelationships.Count(r => r != null && r.IsChecked);
            BtnAddPair.Visibility = (tickedCount == 1)
                ? System.Windows.Visibility.Visible
                : System.Windows.Visibility.Collapsed;
        }

        private void Header_MouseDown(object sender, MouseButtonEventArgs e) { if (e.LeftButton == MouseButtonState.Pressed) DragMove(); }
        private void BtnClose_Click(object sender, RoutedEventArgs e) => Close();

        // --- 1. SỰ KIỆN CHỌN DÒNG -> GỬI LỆNH SELECT VỀ MODEL ---
        private void DgReport_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // [CHẶN VÒNG LẶP] Nếu đây là code tự chọn (do Revit gọi sang), thì DỪNG NGAY.
            if (_isProgrammaticSelection) return;

            // Gom tất cả ID của các dòng đang được chọn
            var idsToSelect = new List<ElementId>();

            // DgReport.SelectedItems trả về danh sách các dòng đang bôi đen
            foreach (var obj in DgReport.SelectedItems)
            {
                if (obj is ReportItem item)
                {
                    if (item.ChildId != ElementId.InvalidElementId)
                    {
                        idsToSelect.Add(item.ChildId);
                    }
                    if (item.ParentId != ElementId.InvalidElementId) // Chọn cả Parent cho dễ nhìn
                    {
                        idsToSelect.Add(item.ParentId);
                    }
                }
            }

            // Gửi yêu cầu Select nếu có ít nhất 1 ID
            if (idsToSelect.Count > 0)
            {
                if (_handler != null && _exEvent != null)
                {
                    _handler.ElementsToSelect = idsToSelect;
                    _handler.Request = StructureSyncRequestId.SelectElements;
                    _exEvent.Raise();
                }
            }
        }

        // --- 2. HÀM TỰ ĐỘNG CHỌN DÒNG (KHI CLICK Ở REVIT - HIGHLIGHT NGƯỢC) ---
        public void AutoSelectRow(ICollection<ElementId> selectedIds)
        {
            if (selectedIds == null || selectedIds.Count == 0) return;

            // Chạy trên luồng UI để đảm bảo an toàn
            this.Dispatcher.Invoke(() =>
            {
                try
                {
                    // [BẬT KHÓA AN TOÀN] Báo hiệu đây là Code đang chọn
                    _isProgrammaticSelection = true;

                    // Ép kiểu DataContext của DataGrid về List<ReportItem> hoặc dùng ItemsSource
                    var items = DgReport.ItemsSource as IEnumerable<ReportItem>;
                    if (items == null) return;

                    // Tìm dòng khớp với ID (Parent hoặc Child)
                    var match = items.FirstOrDefault(item =>
                        selectedIds.Contains(item.ChildId) ||
                        selectedIds.Contains(item.ParentId));

                    if (match != null)
                    {
                        // Nếu dòng đó chưa được chọn thì mới chọn
                        if (!DgReport.SelectedItems.Contains(match))
                        {
                            // Việc gán này sẽ kích hoạt sự kiện SelectionChanged,
                            // nhưng nhờ biến _isProgrammaticSelection nên nó sẽ bị chặn lại -> Hết lỗi Fatal Error
                            DgReport.SelectedItem = match;
                            DgReport.ScrollIntoView(match); // Tự động cuộn xuống dòng đó
                        }
                    }
                }
                catch { }
                finally
                {
                    // [TẮT KHÓA AN TOÀN] Trả lại quyền cho User
                    _isProgrammaticSelection = false;
                }
            });
        }


        // --- 3. CÁC NÚT BẤM CHỨC NĂNG ---

        private void DgReport_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Delete)
            {
                DeleteSelectedRows();
                e.Handled = true; // Chặn hành vi xóa mặc định của DataGrid để tránh lỗi
            }
        }

        private void BtnRemovePair_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.DataContext is ReportItem item)
            {
                // Nếu bấm nút thì chỉ bôi đen dòng đó và gọi hàm xóa
                DgReport.SelectedItem = item;
                DeleteSelectedRows();
            }
        }

        private void DeleteSelectedRows()
        {
            // Lấy danh sách các dòng đang được bôi đen
            var selectedItems = DgReport.SelectedItems.Cast<ReportItem>().ToList();
            if (selectedItems.Count == 0) return;

            if (MessageBox.Show($"Remove {selectedItems.Count} selected pair(s)?", "Confirm", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                var list = (List<ReportItem>)DgReport.ItemsSource;

                // Xóa lần lượt từng Item
                foreach (var item in selectedItems)
                {
                    if (item.SourceItem != null)
                    {
                        // Tìm đúng Index hiện tại của ChildId trong danh sách gốc (rất quan trọng khi xóa nhiều)
                        int idxC = item.SourceItem.ChildIds.IndexOf(item.ChildId);

                        if (idxC != -1)
                        {
                            // 1. Xóa ID khỏi danh sách Cha & Con
                            item.SourceItem.ParentIds.RemoveAt(idxC);
                            item.SourceItem.ChildIds.RemoveAt(idxC);
                            item.SourceItem.ChildCount = item.SourceItem.ChildIds.Count;

                            // 2. QUAN TRỌNG: Xóa luôn phần tử tương ứng trong mảng Offsets
                            if (item.SourceItem.Offsets != null && idxC < item.SourceItem.Offsets.Count)
                            {
                                item.SourceItem.Offsets.RemoveAt(idxC);
                            }
                        }
                    }

                    // Xóa dòng đó khỏi giao diện
                    list.Remove(item);
                }

                // Regenerate report to ensure DataGrid reflects the updated sets.
                RefreshReports();
            }
        }

        public void RefreshDataGrid()
        {
            DgReport.Items.Refresh();
        }

        private List<BuiltInCategory> GetCategoriesForSyncType(SyncType type)
        {
            switch (type)
            {
                case SyncType.PileCap:
                case SyncType.Pile:
                    return new List<BuiltInCategory> { BuiltInCategory.OST_StructuralFoundation };
                case SyncType.Column:
                    return new List<BuiltInCategory> { BuiltInCategory.OST_StructuralColumns };
                case SyncType.DropPanel:
                    return new List<BuiltInCategory> { BuiltInCategory.OST_StructuralFraming };
                case SyncType.Floor:
                    return new List<BuiltInCategory> { BuiltInCategory.OST_Floors };
                case SyncType.Wall:
                    return new List<BuiltInCategory> { BuiltInCategory.OST_Walls };
                default:
                    return new List<BuiltInCategory>();
            }
        }

        private ElementId ExtractPickedParentId(Reference reference)
        {
            if (reference == null) return ElementId.InvalidElementId;
            return (reference.LinkedElementId != ElementId.InvalidElementId) ? reference.LinkedElementId : reference.ElementId;
        }

        private void EnsureOffsetsSize(RelationshipItem rel)
        {
            if (rel == null) return;
            if (rel.Offsets == null) rel.Offsets = new List<string>();
            while (rel.Offsets.Count < rel.ChildIds.Count) rel.Offsets.Add("");
        }

        private void RefreshReports()
        {
            if (_handler?.Logic == null || _sourceRelationships == null) return;
            var newReports = _handler.Logic.GenerateReports(_sourceRelationships, _handler.Links);
            DgReport.ItemsSource = newReports;
            DgReport.Items.Refresh();
        }

        // --- Add new pair into the only-ticked set ---
        private void BtnAddPair_Click(object sender, RoutedEventArgs e)
        {
            if (_handler?.Logic == null || _sourceRelationships == null) return;
            var ticked = _sourceRelationships.Where(r => r != null && r.IsChecked).ToList();
            if (ticked.Count != 1) return;

            var rel = ticked[0];
            if (rel == null) return;

            if (!Enum.TryParse(rel.ParentTypeStr, out SyncType pType)) return;
            if (!Enum.TryParse(rel.ChildTypeStr, out SyncType cType)) return;

            var parentCats = GetCategoriesForSyncType(pType);
            var childCats = GetCategoriesForSyncType(cType);
            bool useLinkForParent = rel.ParentIsFromLink;

            while (true)
            {
                ElementId newParentId = ElementId.InvalidElementId;
                ElementId newChildId = ElementId.InvalidElementId;

                try
                {
                    // 1) Pick Parent
                    this.Hide();
                    var pickedParents = _handler.Logic.SelectParents(parentCats, useLinkForParent);
                    this.Show();

                    if (pickedParents == null || pickedParents.Count == 0) break;
                    newParentId = ExtractPickedParentId(pickedParents[0]);
                    if (newParentId == ElementId.InvalidElementId) break;

                    // 2) Pick Child (always local)
                    this.Hide();
                    var pickedChildren = _handler.Logic.SelectChildren(childCats);
                    this.Show();

                    if (pickedChildren == null || pickedChildren.Count == 0) break;
                    newChildId = pickedChildren[0];
                }
                catch (Exception ex)
                {
                    try { this.Show(); } catch { }
                    MessageBox.Show("Add pair failed: " + ex.Message);
                    break;
                }

                EnsureOffsetsSize(rel);
                string offsetStr = _handler.Logic.ComputeOffsetStringForPair(newParentId, newChildId, _handler.Links);
                if (string.IsNullOrEmpty(offsetStr))
                {
                    MessageBox.Show(
                            "Unable to compute offset for the selected Parent/Child pair.\n\n" +
                            "Most commonly: the Parent was picked from a link, but that link is not ticked in `links` for the current run.\n" +
                            "Please select the correct Parent from the links you are syncing so the row is no longer marked as `Missing Parent`.",
                            "Offset compute failed",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);
                    break;
                }

                rel.ParentIds.Add(newParentId);
                rel.ChildIds.Add(newChildId);
                rel.Offsets.Add(offsetStr);
                rel.ChildCount = rel.ChildIds.Count;

                RefreshReports();

                if (MessageBox.Show("Add another pair to this set?", "Add Pair", MessageBoxButton.YesNo) != MessageBoxResult.Yes)
                    break;
            }
        }

        // --- Edit missing Parent/Child (supports multi-selection) ---
        private void BtnEditPair_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button btn && btn.DataContext is ReportItem clickedItem)) return;
            if (clickedItem == null || clickedItem.SourceItem == null) return;

            // Multi-selection behavior:
            // - If clicked row is already selected -> edit all selected rows.
            // - Otherwise -> edit only clicked row.
            var selectedItems = DgReport.SelectedItems.Cast<ReportItem>().ToList();
            var itemsToProcess = selectedItems.Contains(clickedItem)
                ? selectedItems
                : new List<ReportItem> { clickedItem };

            itemsToProcess = itemsToProcess
                .Where(x => x != null && x.SourceItem != null &&
                            (x.StatusDisplay == "Missing Parent" || x.StatusDisplay == "Missing Child"))
                .ToList();

            if (itemsToProcess.Count == 0) return;

            try
            {
                // Phase 1: edit all Missing Parent rows (pick Parents once)
                var parentItems = itemsToProcess.Where(x => x.StatusDisplay == "Missing Parent").ToList();
                if (parentItems.Count > 0)
                {
                    var firstRel = parentItems[0].SourceItem;
                    if (!Enum.TryParse(firstRel.ParentTypeStr, out SyncType pType)) return;

                    var parentCats = GetCategoriesForSyncType(pType);
                    bool useLinkForParent = firstRel.ParentIsFromLink;

                    this.Hide();
                    var pickedParents = _handler.Logic.SelectParents(parentCats, useLinkForParent);
                    this.Show();

                    if (pickedParents == null || pickedParents.Count < parentItems.Count)
                    {
                        MessageBox.Show(
                            "Please pick at least the same number of Parents as the number of selected rows that are `Missing Parent`.",
                            "Not enough Parents picked",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);
                        return;
                    }

                    for (int i = 0; i < parentItems.Count; i++)
                    {
                        var item = parentItems[i];
                        var rel = item.SourceItem;
                        int idx = item.IndexInSet;
                        if (rel == null) continue;
                        if (idx < 0 || idx >= rel.ChildIds.Count || idx >= rel.ParentIds.Count) continue;

                        var newParentId = ExtractPickedParentId(pickedParents[i]);
                        if (newParentId == ElementId.InvalidElementId) return;

                        rel.ParentIds[idx] = newParentId;
                        EnsureOffsetsSize(rel);

                        string offsetStr = _handler.Logic.ComputeOffsetStringForPair(rel.ParentIds[idx], rel.ChildIds[idx], _handler.Links);
                        if (string.IsNullOrEmpty(offsetStr))
                        {
                            MessageBox.Show(
                                "Unable to compute offset for the newly selected Parent.\n\n" +
                                "The Parent may come from a link that is not ticked in `links` for the current run.\n" +
                                "Please select the Parent again from the links you are syncing.",
                                "Offset compute failed",
                                MessageBoxButton.OK,
                                MessageBoxImage.Warning);
                            return;
                        }

                        rel.Offsets[idx] = offsetStr;
                        rel.ChildCount = rel.ChildIds.Count;
                    }
                }

                // Phase 2: edit all Missing Child rows (pick Children once)
                var childItems = itemsToProcess.Where(x => x.StatusDisplay == "Missing Child").ToList();
                if (childItems.Count > 0)
                {
                    var firstRel = childItems[0].SourceItem;
                    if (!Enum.TryParse(firstRel.ChildTypeStr, out SyncType cType)) return;

                    var childCats = GetCategoriesForSyncType(cType);

                    this.Hide();
                    var pickedChildren = _handler.Logic.SelectChildren(childCats);
                    this.Show();

                    if (pickedChildren == null || pickedChildren.Count < childItems.Count)
                    {
                        MessageBox.Show(
                            "Please pick at least the same number of Children as the number of selected rows that are `Missing Child`.",
                            "Not enough Children picked",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);
                        return;
                    }

                    for (int i = 0; i < childItems.Count; i++)
                    {
                        var item = childItems[i];
                        var rel = item.SourceItem;
                        int idx = item.IndexInSet;
                        if (rel == null) continue;
                        if (idx < 0 || idx >= rel.ChildIds.Count || idx >= rel.ParentIds.Count) continue;

                        var newChildId = pickedChildren[i];
                        rel.ChildIds[idx] = newChildId;
                        EnsureOffsetsSize(rel);

                        string offsetStr = _handler.Logic.ComputeOffsetStringForPair(rel.ParentIds[idx], rel.ChildIds[idx], _handler.Links);
                        if (string.IsNullOrEmpty(offsetStr))
                        {
                            MessageBox.Show(
                                "Unable to compute offset for the newly selected Child.\n\n" +
                                "Please select the correct Child again (Child is selected locally in the current workflow).",
                                "Offset compute failed",
                                MessageBoxButton.OK,
                                MessageBoxImage.Warning);
                            return;
                        }

                        rel.Offsets[idx] = offsetStr;
                        rel.ChildCount = rel.ChildIds.Count;
                    }
                }
            }
            catch (Exception ex)
            {
                try { this.Show(); } catch { }
                MessageBox.Show("Edit pair failed: " + ex.Message);
                return;
            }

            RefreshReports();
        }

        private void BtnSyncPair_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.DataContext is ReportItem currentItem)
            {
                List<ReportItem> itemsToProcess = new List<ReportItem>();

                // 1. Lấy danh sách các dòng đang được bôi đen (Multi-selection)
                var selectedItems = DgReport.SelectedItems.Cast<ReportItem>().ToList();

                // 2. Kiểm tra logic chọn
                // Nếu người dùng bấm nút Sync trên một dòng ĐANG ĐƯỢC CHỌN -> Sync hết cả đám đã chọn
                if (selectedItems.Contains(currentItem))
                {
                    itemsToProcess = selectedItems;
                }
                else
                {
                    // Nếu bấm vào dòng không được chọn (ít xảy ra), chỉ Sync dòng đó
                    itemsToProcess = new List<ReportItem> { currentItem };
                }

                // 3. Lọc bỏ những dòng đã Synced (màu xanh) để tránh chạy thừa
                itemsToProcess = itemsToProcess.Where(x => x.StatusDisplay != "Synced").ToList();

                if (itemsToProcess.Count == 0) return;

                // 4. Gửi danh sách đi xử lý
                _handler.ItemsToSync = itemsToProcess;
                _handler.ReportWindow = this;
                _handler.Request = StructureSyncRequestId.SyncSpecific;

                // Kích hoạt External Event
                _exEvent.Raise();
            }
        }
    }
}
