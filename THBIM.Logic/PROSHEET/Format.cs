using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using DB = Autodesk.Revit.DB;

namespace THBIM
{
    public partial class THBIMSheetWindow
    {
        private void InitFormatTab()
        {
            LoadDwgExportSetups();
        }

        private void TabFormat_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (!_masterList.Any(x => x.IsSelected))
            {
                e.Handled = true;
                MessageBox.Show("Please select sheets in the Selection tab first!", "Notice", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void UpdateCombineList()
        {
            // Kiểm tra tránh lỗi Null
            if (CombineOrderList == null)
            {
                CombineOrderList = new System.Collections.ObjectModel.ObservableCollection<SheetItem>();
            }

            var selectedItems = _masterList.Where(x => x.IsSelected).ToList();

            // 1. Xóa đi những bản vẽ mà người dùng vừa bỏ tick ở Tab Selection
            var toRemove = CombineOrderList.Where(x => !x.IsSelected).ToList();
            foreach (var item in toRemove)
            {
                CombineOrderList.Remove(item);
            }

            // 2. Thêm những bản vẽ MỚI ĐƯỢC TICK vào cuối danh sách (không làm xáo trộn thứ tự cũ)
            foreach (var item in selectedItems)
            {
                if (!CombineOrderList.Contains(item))
                {
                    CombineOrderList.Add(item);
                }
            }

            // ĐÃ XÓA ĐOẠN GỌI LstOrder Ở ĐÂY VÌ MÌNH KHÔNG CÒN DÙNG NỮA
        }





        private void LoadDwgExportSetups()
        {
            // 1. Quét toàn bộ DWG Export Setup có trong Project hiện tại
            var setups = new DB.FilteredElementCollector(_doc)
                .OfClass(typeof(DB.ExportDWGSettings))
                .Cast<DB.ExportDWGSettings>()
                .Select(s => new ExportSetupItem { Id = s.Id, Name = s.Name })
                .OrderBy(s => s.Name)
                .ToList();

            // 2. Thêm một tùy chọn "In-Session Export Setup" ảo lên dòng đầu tiên
            // Dùng InvalidElementId để đánh dấu đây là phiên cài đặt tức thời (chưa lưu)
            setups.Insert(0, new ExportSetupItem
            {
                Id = DB.ElementId.InvalidElementId,
                Name = "<In-Session Export Setup>"
            });

            // 3. Gán danh sách này vào ComboBox trên giao diện
            CboDwgExportSetup.ItemsSource = setups;

            // 4. Mặc định chọn dòng đầu tiên
            if (setups.Any())
            {
                CboDwgExportSetup.SelectedIndex = 0;
            }
        }

        // Giả sử bạn gắn sự kiện Click này vào nút "Export" ở Tab Create hoặc Format
        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            var selectedIds = _masterList.Where(x => x.IsSelected).Select(x => x.Id).ToList();
            if (!selectedIds.Any())
            {
                MessageBox.Show("No sheets selected for export!");
                return;
            }

            _handler.Req.Type = RequestType.Export;
            _handler.Req.Ids = selectedIds;

            // 1. Placement (Đổi từ mm trên giao diện sang Feet của Revit: 1 Feet = 304.8 mm)
            _handler.Req.PaperPlacement = (RbOffset.IsChecked == true) ? DB.PaperPlacementType.LowerLeft : DB.PaperPlacementType.Center;
            double.TryParse(TxtOffsetX.Text, out double ox);
            double.TryParse(TxtOffsetY.Text, out double oy);
            _handler.Req.OffsetX = ox / 304.8;
            _handler.Req.OffsetY = oy / 304.8;

            // 2. Appearance (Ép kiểu chỉ mục ComboBox sang Enum của Revit API)
            _handler.Req.RasterQuality = CboRasterQuality.SelectedIndex switch
            {
                0 => DB.RasterQualityType.Low,
                1 => DB.RasterQualityType.Medium,
                3 => DB.RasterQualityType.Presentation,
                _ => DB.RasterQualityType.High // Mặc định là High (Index 2)
            };

            _handler.Req.ColorDepth = CboColors.SelectedIndex switch
            {
                1 => DB.ColorDepthType.BlackLine,
                2 => DB.ColorDepthType.GrayScale,
                _ => DB.ColorDepthType.Color // Mặc định Color (Index 0)
            };

            // 3. Options
            _handler.Req.HideRefPlanes = ChkHideRefPlanes.IsChecked == true;
            _handler.Req.HideUnrefTags = ChkHideUnrefTags.IsChecked == true;
            _handler.Req.HideCrop = ChkHideCrop.IsChecked == true;
            _handler.Req.HideScopeBox = ChkHideScopeBox.IsChecked == true;

            // 4. Grouping
            _handler.Req.IsCombine = RbCombine.IsChecked == true;
            _handler.Req.CombinedName = TxtCombinedName.Text;

            // GỌI REVIT CHẠY!
            _exEvent.Raise();
        }

        private void ExecuteExport()
        {
            var selectedItems = _masterList.Where(x => x.IsSelected).ToList();
            if (!selectedItems.Any())
            {
                MessageBox.Show("No sheets selected for export!");
                return;
            }

            _handler.Req.Type = RequestType.Export;
            _handler.Req.Ids = selectedItems.Select(x => x.Id).ToList();

            // Đọc trạng thái Tích chọn định dạng
            _handler.Req.ExportPDF = ChkExportPDF.IsChecked == true;
            _handler.Req.ExportDWG = ChkExportDWG.IsChecked == true;


            // QUAN TRỌNG: Truyền danh sách tên Custom vào biến Dictionary
            _handler.Req.CustomNames.Clear();
            foreach (var item in selectedItems)
            {
                _handler.Req.CustomNames[item.Id] = item.CustomFileName;
            }

            // Truyền các cài đặt Checkbox/Radio vào
            _handler.Req.IsCombine = RbCombine.IsChecked == true;
            _handler.Req.CombinedName = TxtCombinedName.Text;

            _handler.Req.PaperPlacement = (RbOffset.IsChecked == true) ? DB.PaperPlacementType.LowerLeft : DB.PaperPlacementType.Center;
            double.TryParse(TxtOffsetX.Text, out double ox);
            double.TryParse(TxtOffsetY.Text, out double oy);
            _handler.Req.OffsetX = ox / 304.8;
            _handler.Req.OffsetY = oy / 304.8;

            _handler.Req.RasterQuality = CboRasterQuality.SelectedIndex switch
            {
                0 => DB.RasterQualityType.Low,
                1 => DB.RasterQualityType.Medium,
                3 => DB.RasterQualityType.Presentation,
                _ => DB.RasterQualityType.High
            };

            _handler.Req.ColorDepth = CboColors.SelectedIndex switch
            {
                1 => DB.ColorDepthType.BlackLine,
                2 => DB.ColorDepthType.GrayScale,
                _ => DB.ColorDepthType.Color
            };

            _handler.Req.HideRefPlanes = ChkHideRefPlanes.IsChecked == true;
            _handler.Req.HideUnrefTags = ChkHideUnrefTags.IsChecked == true;
            _handler.Req.HideCrop = ChkHideCrop.IsChecked == true;
            _handler.Req.HideScopeBox = ChkHideScopeBox.IsChecked == true;

            // THU THẬP DỮ LIỆU DWG
            if (CboDwgExportSetup.SelectedItem is ExportSetupItem setupItem)
            {
                _handler.Req.DwgSetupId = setupItem.Id;
            }
            _handler.Req.DwgExportAsXref = ChkExportViewsOnSheets.IsChecked == true;
            _handler.Req.DwgCleanPcp = ChkCleanPcp.IsChecked == true;
            _handler.Req.DwgBindImages = ChkBindImages.IsChecked == true;

            // Truyền Custom Name
            _handler.Req.CustomNames.Clear();
            foreach (var item in selectedItems) _handler.Req.CustomNames[item.Id] = item.CustomFileName;





            // KÍCH HOẠT REVIT
            _exEvent.Raise();
        }
    }
}