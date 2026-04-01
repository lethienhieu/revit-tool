using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using DB = Autodesk.Revit.DB;


namespace THBIM
{
    public class RequestHandler : IExternalEventHandler
    {
        public RequestData Req { get; set; } = new RequestData();
        public Action OnRequestCompleted;
        public Action<int, int> OnProgressUpdated { get; set; }




        public void Execute(UIApplication app)
        {
            var doc = app.ActiveUIDocument.Document;

            try
            {
                // =======================================================
                // SỬA LỖI FATAL ERROR: Tách lệnh Export ra ngoài Transaction
                // =======================================================
                if (Req.Type == RequestType.Export)
                {
                    DoExport(doc);
                    // Tự động bung màn hình thư mục chứa file vừa xuất
                    if (System.IO.Directory.Exists(Req.FolderPath))
                    {
                        System.Diagnostics.Process.Start("explorer.exe", Req.FolderPath);
                    }
                    OnRequestCompleted?.Invoke();
                    return; // THOÁT LUÔN ĐỂ KHÔNG CHẠY VÀO TRANSACTION BÊN DƯỚI
                }

                // =======================================================
                // CHỈ mở Transaction cho các lệnh chỉnh sửa Database (Tạo/Xóa Set)
                // =======================================================
                using (var t = new DB.Transaction(doc, "THBIM Process"))
                {
                    t.Start();

                    if (Req.Type == RequestType.CreateSet) CreateNewSet(doc, Req.SetName, Req.Ids);
                    else if (Req.Type == RequestType.AddToSet) UpdateExistingSet(doc, Req.SetId, Req.Ids);
                    else if (Req.Type == RequestType.DeleteSet) doc.Delete(Req.SetId);

                    t.Commit();
                }

                OnRequestCompleted?.Invoke();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error API: " + ex.Message, "THBIM Error");
            }


        }



        private void CreateNewSet(DB.Document doc, string name, List<DB.ElementId> ids)
        {
            var pm = doc.PrintManager;
            pm.PrintRange = DB.PrintRange.Select;
            var vss = pm.ViewSheetSetting;

            // Xóa Set cũ nếu bị trùng tên
            var existing = new DB.FilteredElementCollector(doc)
                .OfClass(typeof(DB.ViewSheetSet)).Cast<DB.ViewSheetSet>()
                .FirstOrDefault(x => x.Name == name);

            if (existing != null)
            {
                doc.Delete(existing.Id);
                doc.Regenerate(); // Quan trọng: Cần Regenerate để Revit nhận diện là đã xóa
            }

            // Khai báo biến newViewSet để chứa danh sách bản vẽ
            var newViewSet = new DB.ViewSet();
            foreach (var id in ids)
            {
                var v = doc.GetElement(id) as DB.View;
                if (v != null) newViewSet.Insert(v);
            }

            try
            {
                vss.CurrentViewSheetSet = vss.InSession; // Trỏ tới bộ nhớ tạm
                vss.InSession.Views = newViewSet; // Ghi đè danh sách mới vào bộ nhớ tạm
                vss.SaveAs(name); // Lưu từ bộ nhớ tạm thành tên mới
            }
            catch { /* Bỏ qua lỗi nếu có */ }
        }

        private void UpdateExistingSet(DB.Document doc, DB.ElementId setId, List<DB.ElementId> newIds)
        {
            var existingSet = doc.GetElement(setId) as DB.ViewSheetSet;
            if (existingSet == null) return;

            string setName = existingSet.Name; // Lưu lại tên trước khi xử lý

            var pm = doc.PrintManager;
            pm.PrintRange = DB.PrintRange.Select;
            var vss = pm.ViewSheetSetting;

            var newViewSet = new DB.ViewSet();

            // Giữ lại các bản vẽ cũ
            foreach (DB.View v in existingSet.Views) newViewSet.Insert(v);

            // Thêm các bản vẽ mới
            foreach (var id in newIds)
            {
                var v = doc.GetElement(id) as DB.View;
                if (v != null && !newViewSet.Contains(v)) newViewSet.Insert(v);
            }

            // =========================================================
            // SỬA LỖI Ở ĐÂY: Lưu vòng vèo qua InSession
            // =========================================================
            try
            {
                vss.CurrentViewSheetSet = vss.InSession; // Trỏ tới bộ nhớ tạm
                vss.InSession.Views = newViewSet; // Bơm toàn bộ list mới (gồm cũ + mới) vào

                // Xóa Set cũ đi
                doc.Delete(setId);
                doc.Regenerate();

                // Lưu list mới này với cái tên vừa xóa
                vss.SaveAs(setName);
            }
            catch { /* Bỏ qua lỗi nếu có */ }
        }

        // =========================================================
        // BỘ CHUYỂN ĐỔI KHỔ GIẤY (Chữa lỗi Size giấy không ăn)
        // =========================================================
        private DB.ExportPaperFormat GetPaperFormat(string size)
        {
            if (string.IsNullOrEmpty(size)) return DB.ExportPaperFormat.Default;
            switch (size.ToUpper().Trim())
            {
                // Nhóm ISO A
                case "A0": return DB.ExportPaperFormat.ISO_A0;
                case "A1": return DB.ExportPaperFormat.ISO_A1;
                case "A2": return DB.ExportPaperFormat.ISO_A2;
                case "A3": return DB.ExportPaperFormat.ISO_A3;
                case "A4": return DB.ExportPaperFormat.ISO_A4;

                // Nhóm ARCH
                case "ARCH A": return DB.ExportPaperFormat.ARCH_A;
                case "ARCH B": return DB.ExportPaperFormat.ARCH_B;
                case "ARCH C": return DB.ExportPaperFormat.ARCH_C;
                case "ARCH D": return DB.ExportPaperFormat.ARCH_D;
                case "ARCH E": return DB.ExportPaperFormat.ARCH_E;
                case "ARCH E1": return DB.ExportPaperFormat.ARCH_E1;

                // Nhóm ANSI
                case "ANSI A": return DB.ExportPaperFormat.ANSI_A;
                case "ANSI B": return DB.ExportPaperFormat.ANSI_B;
                case "ANSI C": return DB.ExportPaperFormat.ANSI_C;
                case "ANSI D": return DB.ExportPaperFormat.ANSI_D;
                case "ANSI E": return DB.ExportPaperFormat.ANSI_E;

                // Với các khổ ISO B0, B1... hoặc Custom, Revit API không hỗ trợ sẵn. 
                // Ta trả về Default để Revit tự động quét và lấy kích thước chuẩn của Khung Tên!
                default: return DB.ExportPaperFormat.Default;
            }
        }

        private DB.PageOrientationType GetOrientation(string orientation)
        {
            if (orientation == "Portrait") return DB.PageOrientationType.Portrait;
            return DB.PageOrientationType.Landscape;
        }

        // =========================================================
        // TRUNG TÂM ĐIỀU PHỐI EXPORT 
        // =========================================================
        private void DoExport(DB.Document doc)
        {
            string baseFolder = Req.ExportFolder;
            if (string.IsNullOrWhiteSpace(baseFolder))
            {
                baseFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            }

            string pdfFolder = baseFolder;
            string dwgFolder = baseFolder;

            if (Req.SaveSplitFormat)
            {
                pdfFolder = System.IO.Path.Combine(baseFolder, "PDF");
                dwgFolder = System.IO.Path.Combine(baseFolder, "DWG");

                if (!System.IO.Directory.Exists(pdfFolder)) System.IO.Directory.CreateDirectory(pdfFolder);
                if (!System.IO.Directory.Exists(dwgFolder)) System.IO.Directory.CreateDirectory(dwgFolder);
            }
            else
            {
                if (!System.IO.Directory.Exists(baseFolder)) System.IO.Directory.CreateDirectory(baseFolder);
            }

            // Dùng trực tiếp danh sách Tasks để giữ nguyên thứ tự sắp xếp (Combine Order)
            var pdfTasks = Req.Tasks.Where(x => x.ExportFormat == "PDF").ToList();
            var dwgTasks = Req.Tasks.Where(x => x.ExportFormat == "DWG").ToList();

            if (pdfTasks.Any())
            {
                DoExportPDF(doc, pdfFolder, pdfTasks);
            }

            if (dwgTasks.Any())
            {
                DoExportDWG(doc, dwgFolder, dwgTasks);
            }
        }

        // =========================================================
        // HÀM XUẤT PDF
        // =========================================================
        // =========================================================
        // HÀM XUẤT PDF (SỬ DỤNG MẶC ĐỊNH CỦA REVIT)
        // =========================================================
        private void DoExportPDF(DB.Document doc, string folderPath, List<ExportTaskItem> pdfTasks)
        {
            DB.PDFExportOptions opt = new DB.PDFExportOptions();

            // 1. ÁP DỤNG CÁC CÀI ĐẶT CHUNG TỪ TAB FORMAT
            opt.PaperPlacement = Req.PaperPlacement;
            if (opt.PaperPlacement == DB.PaperPlacementType.LowerLeft)
            {
                opt.OriginOffsetX = Req.OffsetX;
                opt.OriginOffsetY = Req.OffsetY;
            }

            opt.RasterQuality = Req.RasterQuality;
            opt.ColorDepth = Req.ColorDepth;
            opt.HideReferencePlane = Req.HideRefPlanes;
            opt.HideUnreferencedViewTags = Req.HideUnrefTags;
            opt.HideCropBoundaries = Req.HideCrop;
            opt.HideScopeBoxes = Req.HideScopeBox;

            // 2. LOGIC TỶ LỆ IN (CHỐNG BỊ CẮT LỀ BẢN VẼ)
            if (Req.UseFitToPage)
            {
                opt.ZoomType = DB.ZoomType.FitToPage;
            }
            else
            {
                opt.ZoomType = DB.ZoomType.Zoom;
                opt.ZoomPercentage = Req.ZoomPercentage;
            }


            // =========================================================
            // NHÁNH 1: XUẤT GỘP (COMBINE BẰNG REVIT API)
            // =========================================================
            if (Req.IsCombine)
            {
                try
                {
                    foreach (var task in pdfTasks) UpdateUI(task, "Publishing...");

                    opt.Combine = true;

                    string finalCombinedName = string.IsNullOrWhiteSpace(Req.CombinedName) ? "THBIM_Combined" : Req.CombinedName;
                    char[] invalidChars = System.IO.Path.GetInvalidFileNameChars();
                    foreach (char c in invalidChars) finalCombinedName = finalCombinedName.Replace(c.ToString(), "_");

                    opt.FileName = finalCombinedName;

                    // --- KIỂM TRA QUYẾT ĐỊNH TỪ GIAO DIỆN (TAB CREATE) ---
                    // Lấy danh sách các khổ giấy người dùng đã chọn trên giao diện
                    var distinctSizes = pdfTasks.Select(x => x.PaperSize).Distinct().ToList();

                    // TRƯỜNG HỢP A: Có nhiều khổ giấy khác nhau trộn lẫn (VD: A0 và A1)
                    if (distinctSizes.Count > 1)
                    {
                        // Bắt buộc dùng Default để file PDF tạo ra có nhiều trang với kích thước to nhỏ khác nhau
                        opt.PaperFormat = DB.ExportPaperFormat.Default;
                        opt.ZoomType = DB.ZoomType.Zoom;
                        opt.ZoomPercentage = 100;
                    }
                    // TRƯỜNG HỢP B: Người dùng thiết lập tất cả cùng 1 khổ giấy (VD: Đều là A1)
                    else
                    {
                        var firstTask = pdfTasks.FirstOrDefault();
                        if (firstTask != null)
                        {
                            // Tôn trọng 100% khổ giấy và hướng giấy trên giao diện
                            opt.PaperFormat = GetPaperFormat(firstTask.PaperSize);
                            opt.PaperOrientation = GetOrientation(firstTask.Orientation);

                            // Tôn trọng tỷ lệ người dùng chọn ở Tab Format
                            if (Req.UseFitToPage)
                            {
                                opt.ZoomType = DB.ZoomType.FitToPage;
                            }
                            else
                            {
                                opt.ZoomType = DB.ZoomType.Zoom;
                                opt.ZoomPercentage = Req.ZoomPercentage;
                            }
                        }
                    }
                    // -----------------------------------------------------

                    List<DB.ElementId> orderedIds = pdfTasks.Select(x => x.Id).ToList();
                    doc.Export(folderPath, orderedIds, opt);

                    foreach (var task in pdfTasks) UpdateUI(task, "Success");
                }
                catch (Exception)
                {
                    foreach (var task in pdfTasks) UpdateUI(task, "Error");
                }
            }
            // =========================================================
            // NHÁNH 2: XUẤT RỜI RẠC TỪNG FILE
            // =========================================================
            else
            {
                foreach (var task in pdfTasks)
                {
                    try
                    {
                        UpdateUI(task, "Publishing...");

                        opt.PaperFormat = GetPaperFormat(task.PaperSize);
                        opt.PaperOrientation = GetOrientation(task.Orientation);
                        opt.Combine = true; // Ép true để Revit không tự thêm linh tinh vào tên file

                        string fileName = string.IsNullOrWhiteSpace(task.CustomFileName) ? task.SheetNumber : task.CustomFileName;
                        char[] invalidChars = System.IO.Path.GetInvalidFileNameChars();
                        foreach (char c in invalidChars) fileName = fileName.Replace(c.ToString(), "_");

                        opt.FileName = fileName;

                        // Chạy lệnh xuất từng file
                        doc.Export(folderPath, new List<DB.ElementId> { task.Id }, opt);

                        UpdateUI(task, "Success");
                    }
                    catch (Exception)
                    {
                        UpdateUI(task, "Error");
                    }
                }
            }
        }

        // =========================================================
        // HÀM XUẤT DWG VÀ CÁC TÍNH NĂNG DỌN DẸP
        // =========================================================
        private void DoExportDWG(DB.Document doc, string folderPath, List<ExportTaskItem> dwgTasks)
        {
            DB.DWGExportOptions dwgOpt = null;

            // SỬA LỖI FATAL ERROR: Thêm check null và InvalidElementId
            if (Req.DwgSetupId != null && Req.DwgSetupId != DB.ElementId.InvalidElementId)
            {
                var setup = doc.GetElement(Req.DwgSetupId) as DB.ExportDWGSettings;
                if (setup != null) dwgOpt = setup.GetDWGExportOptions();
            }

            if (dwgOpt == null) dwgOpt = new DB.DWGExportOptions();

            dwgOpt.MergedViews = !Req.DwgExportAsXref;

            // LUÔN XUẤT DWG RỜI RẠC
            foreach (var task in dwgTasks)
            {
                string fileName = string.IsNullOrWhiteSpace(task.CustomFileName) ? task.SheetNumber : task.CustomFileName;
                char[] invalidChars = System.IO.Path.GetInvalidFileNameChars();
                foreach (char c in invalidChars) fileName = fileName.Replace(c.ToString(), "_");

                doc.Export(folderPath, fileName, new List<DB.ElementId> { task.Id }, dwgOpt);
                UpdateUI(task, "Success");
            }

            if (Req.DwgCleanPcp)
            {
                try
                {
                    string[] pcpFiles = System.IO.Directory.GetFiles(folderPath, "*.pcp");
                    foreach (string file in pcpFiles)
                    {
                        System.IO.File.Delete(file);
                    }
                }
                catch { /* Bỏ qua nếu bị kẹt quyền Windows */ }
            }

            if (Req.DwgBindImages)
            {
                MessageBox.Show("", "", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }


        public string GetName() => "THBIM Handler";

        // Hàm giúp cập nhật trạng thái lên giao diện ngay lập tức
        // Hàm cập nhật giao diện thời gian thực (Đã fix lỗi nhảy đồng loạt)
        // Hàm cập nhật giao diện thời gian thực
        private void UpdateUI(ExportTaskItem task, string status)
        {
            task.ExportStatus = status;

            System.Windows.Application.Current.Dispatcher.Invoke(
                System.Windows.Threading.DispatcherPriority.Background,
                new Action(delegate {

                    // Đếm số lượng file ĐÃ CHẠY XONG (Thành công hoặc Lỗi đều tính là xong)
                    if (Req != null && Req.Tasks != null && OnProgressUpdated != null)
                    {
                        int total = Req.Tasks.Count;
                        int completed = Req.Tasks.Count(x => x.ExportStatus == "Success" || x.ExportStatus == "Error");

                        // Gọi đường dây nóng báo về giao diện
                        OnProgressUpdated.Invoke(completed, total);
                    }

                }));
        }

    }




    // --- DATA MODELS ---
    public class RequestData
    {
        public RequestType Type { get; set; }
        public List<DB.ElementId> Ids { get; set; }
        public string SetName { get; set; }
        public DB.ElementId SetId { get; set; }

        public bool IsCombine { get; set; }
        public string CombinedName { get; set; }
        public DB.PaperPlacementType PaperPlacement { get; set; }
        public double OffsetX { get; set; }
        public double OffsetY { get; set; }
        public DB.RasterQualityType RasterQuality { get; set; }
        public DB.ColorDepthType ColorDepth { get; set; }
        public bool HideRefPlanes { get; set; }
        public bool HideUnrefTags { get; set; }
        public bool HideCrop { get; set; }
        public bool HideScopeBox { get; set; }

        public bool ExportPDF { get; set; }
        public bool ExportDWG { get; set; }

        public bool UseFitToPage { get; set; } = true;
        public int ZoomPercentage { get; set; } = 100;

        // CÁC BIẾN CHO DWG EXPORT
        public DB.ElementId DwgSetupId { get; set; }
        public bool DwgExportAsXref { get; set; }
        public bool DwgCleanPcp { get; set; }
        public bool DwgBindImages { get; set; }
        public string FolderPath { get; set; }

        // THÊM 3 BIẾN NÀY DÀNH CHO TAB CREATE
        public string ExportFolder { get; set; }
        public bool SaveSplitFormat { get; set; }
        public List<ExportTaskItem> Tasks { get; set; } = new List<ExportTaskItem>();

        // BẮT BUỘC THÊM BIẾN NÀY ĐỂ MANG TÊN CUSTOM TỪ GIAO DIỆN XUỐNG
        public Dictionary<DB.ElementId, string> CustomNames { get; set; } = new Dictionary<DB.ElementId, string>();
    }

    public class ExportSetupItem
    {
        public DB.ElementId Id { get; set; }
        public string Name { get; set; }
    }

    public enum RequestType { None, CreateSet, AddToSet, DeleteSet, Export }

    public class SetItem { public DB.ElementId Id { get; set; } public string Name { get; set; } }

    public class SheetItem : INotifyPropertyChanged
    {
        public DB.ElementId Id { get; set; }

        private bool _sel;
        public bool IsSelected
        {
            get => _sel;
            set { _sel = value; OnPropertyChanged(); }
        }

        public string SheetNumber { get; set; }
        public string SheetName { get; set; }
        public string Revision { get; set; }
        public string PaperSize { get; set; }
        public string ViewType { get; set; }
        public string ViewScale { get; set; }
        public string DetailLevel { get; set; }
        public string Discipline { get; set; }

        private string _customFileName;
        public string CustomFileName
        {
            get => _customFileName;
            set { _customFileName = value; OnPropertyChanged(); }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string n = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));
    }

    public class NameRuleItem : INotifyPropertyChanged
    {
        public string ParameterName { get; set; }
        public string SampleValue { get; set; }

        private string _prefix = "";
        public string Prefix { get => _prefix; set { _prefix = value; OnPropertyChanged(); } }

        private string _suffix = "";
        public string Suffix { get => _suffix; set { _suffix = value; OnPropertyChanged(); } }

        private string _separator = "-";
        public string Separator { get => _separator; set { _separator = value; OnPropertyChanged(); } }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public class ExportTaskItem : INotifyPropertyChanged
    {
        public DB.ElementId Id { get; set; }
        public string SheetNumber { get; set; }
        public string SheetName { get; set; }

        public string ExportFormat { get; set; } // PDF hoặc DWG
        public string CustomFileName { get; set; }

        private string _paperSize;
        public string PaperSize
        {
            get => _paperSize;
            set { _paperSize = value; OnPropertyChanged(); }
        }

        private string _orientation = "Landscape";
        public string Orientation
        {
            get => _orientation;
            set { _orientation = value; OnPropertyChanged(); }
        }




        private string _exportStatus = "Pending";
        public string ExportStatus
        {
            get => _exportStatus;
            set { _exportStatus = value; OnPropertyChanged(); }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string n = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));
    }
}