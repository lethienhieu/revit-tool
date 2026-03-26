using Autodesk.Revit.DB;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using DB = Autodesk.Revit.DB;
// ĐÃ XÓA: using System.Windows.Forms; (Không cần nữa trên Revit 2025 / .NET 8)

namespace THBIM
{
    public partial class THBIMSheetWindow
    {
        // Danh sách chuyên dụng cho Tab Create
        public ObservableCollection<ExportTaskItem> CreateTaskList { get; set; } = new ObservableCollection<ExportTaskItem>();

        // Hàm này sẽ được gọi mỗi khi người dùng bấm sang Tab Create
        private void UpdateCreateList()
        {
            if (CreateTaskList == null)
            {
                CreateTaskList = new ObservableCollection<ExportTaskItem>();
            }

            var selectedItems = _masterList.Where(x => x.IsSelected).ToList();
            bool isExportPdf = ChkExportPDF.IsChecked == true;
            bool isExportDwg = ChkExportDWG.IsChecked == true;

            // 1. DỌN DẸP (Xóa những task không còn được chọn hoặc bị bỏ tick định dạng)
            var tasksToRemove = CreateTaskList.Where(task =>
            {
                bool stillSelected = selectedItems.Any(item => item.Id == task.Id);
                bool formatStillChecked = (task.ExportFormat == "PDF" && isExportPdf) || (task.ExportFormat == "DWG" && isExportDwg);
                return !stillSelected || !formatStillChecked;
            }).ToList();

            foreach (var task in tasksToRemove) CreateTaskList.Remove(task);

            // ===============================================================
            // 1.5 CẬP NHẬT ĐỒNG BỘ DỮ LIỆU (Fix lỗi đổi tên không nhận)
            // ===============================================================
            foreach (var task in CreateTaskList)
            {
                var sourceItem = selectedItems.FirstOrDefault(x => x.Id == task.Id);
                if (sourceItem != null)
                {
                    // Liên tục chép đè dữ liệu mới nhất từ Tab Selection sang
                    task.CustomFileName = sourceItem.CustomFileName;
                    // Nếu bạn có sửa PaperSize ở Tab Selection thì cũng thêm vào đây:
                    // task.PaperSize = sourceItem.PaperSize; 
                }
            }

            // 2. THÊM MỚI (Chỉ chạy cho những bản vẽ chưa từng được đưa vào danh sách)
            foreach (var item in selectedItems)
            {
                if (isExportPdf && !CreateTaskList.Any(x => x.Id == item.Id && x.ExportFormat == "PDF"))
                {
                    CreateTaskList.Add(new ExportTaskItem
                    {
                        Id = item.Id,
                        SheetNumber = item.SheetNumber,
                        SheetName = item.SheetName,
                        ExportFormat = "PDF",
                        PaperSize = item.PaperSize,
                        Orientation = "Landscape",
                        ExportStatus = "Pending",
                        CustomFileName = item.CustomFileName // Lấy đúng tên Custom
                    });
                }

                if (isExportDwg && !CreateTaskList.Any(x => x.Id == item.Id && x.ExportFormat == "DWG"))
                {
                    CreateTaskList.Add(new ExportTaskItem
                    {
                        Id = item.Id,
                        SheetNumber = item.SheetNumber,
                        SheetName = item.SheetName,
                        ExportFormat = "DWG",
                        PaperSize = item.PaperSize,
                        Orientation = "Landscape",
                        ExportStatus = "Pending",
                        CustomFileName = item.CustomFileName // Lấy đúng tên Custom
                    });
                }
            }

            // 3. GẮN VÀO GIAO DIỆN
            if (DgCreate.ItemsSource == null) DgCreate.ItemsSource = CreateTaskList;

            // 4. CẬP NHẬT GIAO DIỆN THANH PROGRESS
            TxtOverallProgress.Text = $"Completed 0 of {CreateTaskList.Count} tasks";
            PbOverall.Maximum = CreateTaskList.Count > 0 ? CreateTaskList.Count : 100;
            PbOverall.Value = 0;
        }

        private void BtnBrowseFolder_Click(object sender, RoutedEventArgs e)
        {
#if NET8_0_OR_GREATER
            var dialog = new Microsoft.Win32.OpenFolderDialog
            {
                Title = "Select a folder to export files",
                Multiselect = false
            };
            if (dialog.ShowDialog() == true)
            {
                TxtExportFolder.Text = dialog.FolderName;
            }
#else
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "Select a folder to export files"
            })
            {
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    TxtExportFolder.Text = dialog.SelectedPath;
                }
            }
#endif
        }

        // Hàm kích hoạt tiến trình xuất file
        // Hàm kích hoạt tiến trình xuất file
        // Hàm kích hoạt tiến trình xuất file
        private void StartExportProcess()
        {
            if (CreateTaskList == null || CreateTaskList.Count == 0)
            {
                MessageBox.Show("No sheets selected for export.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(TxtExportFolder.Text))
            {
                MessageBox.Show("Please choose a folder to save your files.", "Notice", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            _handler.Req.FolderPath = TxtExportFolder.Text;
            _handler.Req.Type = RequestType.Export;
            _handler.Req.ExportFolder = TxtExportFolder.Text;
            _handler.Req.SaveSplitFormat = RbSaveSplitFormat.IsChecked == true;



            // ==========================================================
            // 1. SỬA LỖI SAI THỨ TỰ COMBINE: Lọc theo danh sách CombineOrderList
            // ==========================================================
            if (RbCombine.IsChecked == true && CombineOrderList != null && CombineOrderList.Count > 0)
            {
                var orderedTasks = new System.Collections.Generic.List<ExportTaskItem>();
                foreach (var sheet in CombineOrderList)
                {
                    // Lấy task PDF/DWG tương ứng với Sheet đã được xếp thứ tự
                    orderedTasks.AddRange(CreateTaskList.Where(t => t.Id == sheet.Id));
                }
                // Thêm các task còn sót lại (nếu có)
                foreach (var task in CreateTaskList)
                {
                    if (!orderedTasks.Contains(task)) orderedTasks.Add(task);
                }
                _handler.Req.Tasks = orderedTasks;
            }
            else
            {
                _handler.Req.Tasks = CreateTaskList.ToList();
            }


            if (CreateTaskList != null)
            {
                foreach (var task in CreateTaskList)
                {
                    // Đưa chữ về Pending -> XAML sẽ tự động xóa màu nền, trả về ô trắng tinh
                    task.ExportStatus = "Pending";
                }
            }

            // ==========================================================
            // 2. THU THẬP THÔNG SỐ PDF TỪ TAB FORMAT
            // ==========================================================
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
            _handler.Req.IsCombine = RbCombine.IsChecked == true;
            _handler.Req.CombinedName = TxtCombinedName.Text;

            _handler.Req.UseFitToPage = RbFitToPage.IsChecked == true;

            if (int.TryParse(TxtZoomPercentage.Text, out int zoomVal))
                _handler.Req.ZoomPercentage = zoomVal;
            else
                _handler.Req.ZoomPercentage = 100;



            // ==========================================================
            // 3. THU THẬP THÔNG SỐ DWG TỪ TAB FORMAT (SỬA LỖI VĂNG FILE)
            // ==========================================================
            if (CboDwgExportSetup != null && CboDwgExportSetup.SelectedItem is ExportSetupItem setupItem)
            {
                _handler.Req.DwgSetupId = setupItem.Id;
            }
            else
            {
                _handler.Req.DwgSetupId = null; // Báo hiệu là lấy Default
            }

            _handler.Req.DwgExportAsXref = ChkExportViewsOnSheets.IsChecked == true;
            _handler.Req.DwgCleanPcp = ChkCleanPcp.IsChecked == true;
            _handler.Req.DwgBindImages = ChkBindImages.IsChecked == true;

            // Chuyển giao diện
            BtnNext.IsEnabled = false;
            BtnNext.Content = "Create";
            BtnBack.IsEnabled = false;

            // GỌI REVIT CHẠY
            _exEvent.Raise();
        }
    }
}