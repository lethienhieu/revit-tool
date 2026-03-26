using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Input;

namespace THBIM
{
    public class ViewModelBase : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class RelayCommand<T> : ICommand
    {
        private readonly Action<T> _execute;
        private readonly Predicate<T> _canExecute;

        public RelayCommand(Action<T> execute, Predicate<T> canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        public bool CanExecute(object parameter) => _canExecute == null || _canExecute((T)parameter);

        public void Execute(object parameter) => _execute((T)parameter);

        // Kết nối với CommandManager để nút bấm tự động Sáng/Mờ khi dữ liệu thay đổi
        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }
    }

    // ==========================================
    // 1. VIEW MODEL TỔNG QUẢN LÝ (MAIN)
    // ==========================================
    public class MultiCategoryExportViewModel : ViewModelBase
    {
        public Document Doc { get; private set; }
        public ObservableCollection<string> AvailableCategories { get; set; } = new ObservableCollection<string>();
        public ObservableCollection<string> AvailableLevels { get; set; } = new ObservableCollection<string>();

        public ObservableCollection<string> GlobalColumnIds { get; } = new ObservableCollection<string>();
        public ObservableCollection<GrandTotalCellViewModel> GrandTotalCells { get; } = new ObservableCollection<GrandTotalCellViewModel>();

        public ObservableCollection<CategoryTableViewModel> CategoryTables { get; set; } = new ObservableCollection<CategoryTableViewModel>();

        // --- QUẢN LÝ PROFILE VÀ PROJECT INFO ---
        public ProjectInfo CurrentProjectInfo { get; set; } = new ProjectInfo(); // BIẾN LƯU GHI NHỚ PROJECT
        public ObservableCollection<string> SavedProfiles { get; } = new ObservableCollection<string>();
        private Dictionary<string, string> ProfilePaths { get; set; } = new Dictionary<string, string>();

        private string _selectedProfileName;
        public string SelectedProfileName
        {
            get => _selectedProfileName;
            set
            {
                if (_selectedProfileName != value)
                {
                    _selectedProfileName = value;
                    OnPropertyChanged();

                    if (!string.IsNullOrEmpty(_selectedProfileName) && ProfilePaths.ContainsKey(_selectedProfileName))
                    {
                        var profile = ProfileManager.LoadFromFile(ProfilePaths[_selectedProfileName]);
                        if (profile != null)
                        {
                            // 1. Nạp cấu hình. CẢNH BÁO: Quá trình này sẽ TỰ ĐỘNG gọi Revit quét lại mô hình!
                            LoadProfileIntoUI(profile);

                            // =========================================================
                            // BƯỚC 3: KIỂM TRA SỰ THAY ĐỔI VÀ HIỂN THỊ DẠNG BẢNG
                            // =========================================================
                            if (profile.HistorySnapshot != null && profile.HistorySnapshot.Count > 0)
                            {
                                // Tạo một danh sách để chứa các dòng dữ liệu thay đổi
                                var revisionList = new List<RevisionItem>();

                                foreach (var table in CategoryTables)
                                {
                                    foreach (var row in table.Rows)
                                    {
                                        foreach (var cell in row.Cells)
                                        {
                                            if (cell.IsNumeric && cell.Column != null && !string.IsNullOrEmpty(cell.Column.ParameterName))
                                            {
                                                string uniqueKey = $"{table.SelectedCategory}_{table.SelectedLevelValue}_{row.MainParameterValue}_{cell.Column.ParameterName}";
                                                double khoiLuongMoiInternal = cell.NumericValue;

                                                // Dò tìm trong Sổ cũ
                                                if (profile.HistorySnapshot.TryGetValue(uniqueKey, out double khoiLuongCuInternal))
                                                {
                                                    double chenhLechInternal = khoiLuongMoiInternal - khoiLuongCuInternal;

                                                    // Sai số > 0.001 mới tính là có thay đổi
                                                    if (Math.Abs(chenhLechInternal) > 0.001)
                                                    {
                                                        string oldValStr = "0";
                                                        string newValStr = cell.CellValue;
                                                        string diffStr = "0";

                                                        if (cell.ParamDataType != null)
                                                        {
                                                            try
                                                            {
                                                                oldValStr = UnitFormatUtils.Format(Doc.GetUnits(), cell.ParamDataType, khoiLuongCuInternal, false);
                                                                diffStr = UnitFormatUtils.Format(Doc.GetUnits(), cell.ParamDataType, Math.Abs(chenhLechInternal), false);
                                                            }
                                                            catch
                                                            {
                                                                oldValStr = Math.Round(khoiLuongCuInternal, 2).ToString();
                                                                diffStr = Math.Round(Math.Abs(chenhLechInternal), 2).ToString();
                                                            }
                                                        }
                                                        else
                                                        {
                                                            oldValStr = Math.Round(khoiLuongCuInternal, 2).ToString();
                                                            diffStr = Math.Round(Math.Abs(chenhLechInternal), 2).ToString();
                                                        }

                                                        string dau = chenhLechInternal > 0 ? "+" : "-";

                                                        // NẠP VÀO DANH SÁCH BẢNG
                                                        revisionList.Add(new RevisionItem
                                                        {
                                                            Category = table.SelectedCategory,
                                                            Level = table.SelectedLevelValue,
                                                            Element = row.MainParameterValue,
                                                            Parameter = cell.Column.ParameterName,
                                                            OldValue = oldValStr,
                                                            NewValue = newValStr,
                                                            Difference = $"{dau}{diffStr}"
                                                        });
                                                    }
                                                }
                                                else
                                                {
                                                    // NẾU CẤU KIỆN MỚI ĐƯỢC VẼ THÊM
                                                    revisionList.Add(new RevisionItem
                                                    {
                                                        Category = table.SelectedCategory,
                                                        Level = table.SelectedLevelValue,
                                                        Element = row.MainParameterValue,
                                                        Parameter = cell.Column.ParameterName,
                                                        OldValue = "N/A",
                                                        NewValue = cell.CellValue,
                                                        Difference = "NEWLY CREATED"
                                                    });
                                                }
                                            }
                                        }
                                    }
                                }

                                // TỰ ĐỘNG VẼ CỬA SỔ BẢNG WPF (Chỉ bật khi có sự thay đổi)
                                if (revisionList.Count > 0)
                                {
                                    var grid = new System.Windows.Controls.DataGrid
                                    {
                                        ItemsSource = revisionList,
                                        AutoGenerateColumns = true, // Tự động tạo 7 cột dựa theo class RevisionItem
                                        IsReadOnly = true, // Chỉ xem, không cho sửa số
                                        AlternatingRowBackground = System.Windows.Media.Brushes.WhiteSmoke, // Kẻ sọc xám trắng cho dễ nhìn
                                        HeadersVisibility = System.Windows.Controls.DataGridHeadersVisibility.Column,
                                        GridLinesVisibility = System.Windows.Controls.DataGridGridLinesVisibility.All,
                                        Margin = new System.Windows.Thickness(10)
                                    };

                                    var window = new System.Windows.Window
                                    {
                                        Title = $"THBIM - Detected {revisionList.Count} quantity changes!",
                                        Content = grid,
                                        Width = 850,
                                        Height = 500,
                                        WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen,
                                        Topmost = true // Luôn nổi lên trên cùng
                                    };

                                    window.ShowDialog();
                                }
                            }

                        }
                    }
                }
            }
        }

        public class RevisionItem
        {
            public string Category { get; set; }
            public string Level { get; set; }
            public string Element { get; set; }
            public string Parameter { get; set; }
            public string OldValue { get; set; }
            public string NewValue { get; set; }
            public string Difference { get; set; }
        }

        public ICommand AddGlobalColumnCommand { get; set; }
        public ICommand RemoveGlobalColumnCommand { get; set; }
        public ICommand AddCategoryTableCommand { get; set; }
        public ICommand RemoveTableCommand { get; set; }

        // CÁC NÚT QUẢN LÝ XUẤT VÀ LƯU TRỮ
        public ICommand ExportCommand { get; set; }
        public ICommand ImportProfileCommand { get; set; }
        public ICommand SaveProfileCommand { get; set; }
        public ICommand SaveAsProfileCommand { get; set; }
        public ICommand DeleteProfileCommand { get; set; }

        public MultiCategoryExportViewModel(Document doc)
        {
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            Doc = doc;
            LoadActiveRevitData();

            // TẢI TRÍ NHỚ PROFILE CŨ VÀO GIAO DIỆN
            ProfilePaths = ProfileManager.LoadRegistry();
            foreach (var key in ProfilePaths.Keys) SavedProfiles.Add(key);

            AddGlobalColumnCommand = new RelayCommand<object>(o => { string newColId = Guid.NewGuid().ToString(); GlobalColumnIds.Add(newColId); GrandTotalCells.Add(new GrandTotalCellViewModel(newColId)); foreach (var table in CategoryTables) table.AddColumn(newColId); RecalculateGrandTotals(); });
            RemoveGlobalColumnCommand = new RelayCommand<string>(colId => { if (!string.IsNullOrEmpty(colId)) { GlobalColumnIds.Remove(colId); var gtToRemove = GrandTotalCells.FirstOrDefault(g => g.ColumnId == colId); if (gtToRemove != null) GrandTotalCells.Remove(gtToRemove); foreach (var table in CategoryTables) table.RemoveColumn(colId); RecalculateGrandTotals(); } });
            AddCategoryTableCommand = new RelayCommand<object>(o => { var newTable = new CategoryTableViewModel(this); foreach (var id in GlobalColumnIds) newTable.AddColumn(id); CategoryTables.Add(newTable); RecalculateGrandTotals(); });
            RemoveTableCommand = new RelayCommand<CategoryTableViewModel>(table => { if (table != null) { CategoryTables.Remove(table); RecalculateGrandTotals(); } });

            // LOGIC CHO NÚT XUẤT EXCEL
            ExportCommand = new RelayCommand<object>(o =>
            {
                var projWindow = new ProjectInfoWindow(CurrentProjectInfo);
                if (projWindow.ShowDialog() == true)
                {
                    CurrentProjectInfo = projWindow.ResultInfo;

                    // Chỉ bật cửa sổ hỏi nơi LƯU FILE
                    var saveFileDialog = new Microsoft.Win32.SaveFileDialog
                    {
                        Filter = "Excel files (*.xlsx)|*.xlsx",
                        FileName = "Exported_BOQ.xlsx",
                        Title = "Select location to save exported Excel file"
                    };

                    if (saveFileDialog.ShowDialog() == true)
                    {
                        try
                        {
                            var tempProfile = ExtractDataFromUI("ExportTemp");

                            // Trực tiếp gọi hàm xuất file mà không cần templatePath
                            ExcelExportManager.ExportToExcel(tempProfile, saveFileDialog.FileName);

                            // =========================================================
                            // ĐOẠN CODE MỚI: HỎI VÀ MỞ FILE EXCEL NGAY SAU KHI XUẤT XONG
                            // =========================================================
                            var msgResult = System.Windows.MessageBox.Show(
                                "Excel file exported successfully!\n\nDo you want to open the exported Excel file now?",
                                "THBIM QTOPRO",
                                System.Windows.MessageBoxButton.YesNo,
                                System.Windows.MessageBoxImage.Information);

                            if (msgResult == System.Windows.MessageBoxResult.Yes)
                            {
                                // Lệnh gọi Windows mở file
                                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
                                {
                                    FileName = saveFileDialog.FileName,
                                    UseShellExecute = true
                                });
                            }
                            // =========================================================
                        }
                        catch (Exception ex)
                        {
                            System.Windows.MessageBox.Show("Export error: " + ex.Message, "Error");
                        }
                    }
                }
            });

            // LOGIC CHO NÚT IMPORT (.TXT)
            ImportProfileCommand = new RelayCommand<object>(o =>
            {
                var openFileDialog = new Microsoft.Win32.OpenFileDialog { Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*" };
                if (openFileDialog.ShowDialog() == true)
                {
                    var profile = ProfileManager.LoadFromFile(openFileDialog.FileName);
                    if (profile != null)
                    {
                        LoadProfileIntoUI(profile);
                        if (!SavedProfiles.Contains(profile.ProfileName)) SavedProfiles.Add(profile.ProfileName);
                        ProfilePaths[profile.ProfileName] = openFileDialog.FileName;
                        ProfileManager.SaveRegistry(ProfilePaths); // Nhớ vào máy

                        // Cập nhật giá trị mà không kích hoạt event Load đè
                        _selectedProfileName = profile.ProfileName;
                        OnPropertyChanged(nameof(SelectedProfileName));
                    }
                }
            });

            // LOGIC CHO NÚT SAVE AS (Tạo file .TXT mới kèm hỏi thông tin dự án)
            SaveAsProfileCommand = new RelayCommand<object>(o =>
            {
                var projWindow = new ProjectInfoWindow(CurrentProjectInfo);
                if (projWindow.ShowDialog() == true)
                {
                    CurrentProjectInfo = projWindow.ResultInfo; // Cập nhật dự án

                    var saveFileDialog = new Microsoft.Win32.SaveFileDialog { Filter = "Text files (*.txt)|*.txt", FileName = "QTO_Profile_New.txt" };
                    if (saveFileDialog.ShowDialog() == true)
                    {
                        string profileName = System.IO.Path.GetFileNameWithoutExtension(saveFileDialog.FileName);
                        var profileData = ExtractDataFromUI(profileName);

                        ProfileManager.SaveToFile(profileData, saveFileDialog.FileName);

                        if (!SavedProfiles.Contains(profileName)) SavedProfiles.Add(profileName);
                        ProfilePaths[profileName] = saveFileDialog.FileName;
                        ProfileManager.SaveRegistry(ProfilePaths);

                        _selectedProfileName = profileName;
                        OnPropertyChanged(nameof(SelectedProfileName));
                        System.Windows.MessageBox.Show("New Profile created successfully!", "Success");
                    }
                }
            });

            // LOGIC CHO NÚT SAVE (LƯU ĐÈ FILE CŨ)
            SaveProfileCommand = new RelayCommand<object>(
                execute: o =>
                {
                    if (ProfilePaths.TryGetValue(SelectedProfileName, out string path))
                    {
                        var profileData = ExtractDataFromUI(SelectedProfileName);
                        ProfileManager.SaveToFile(profileData, path);
                        System.Windows.MessageBox.Show($"Successfully saved to: {SelectedProfileName}", "Success");
                    }
                    else { SaveAsProfileCommand.Execute(null); }
                },
                canExecute: o => !string.IsNullOrEmpty(SelectedProfileName)
            );

            // LOGIC CHO NÚT DELETE (XÓA FILE & XÓA GIAO DIỆN)
            DeleteProfileCommand = new RelayCommand<object>(
                execute: o =>
                {
                    var result = System.Windows.MessageBox.Show($"Are you sure you want to delete the Profile '{SelectedProfileName}'? The txt file will be permanently deleted.", "Confirmation", System.Windows.MessageBoxButton.YesNo, System.Windows.MessageBoxImage.Warning);
                    if (result == System.Windows.MessageBoxResult.Yes)
                    {
                        string toDelete = SelectedProfileName;
                        if (ProfilePaths.TryGetValue(toDelete, out string path))
                        {
                            if (System.IO.File.Exists(path)) System.IO.File.Delete(path); // Xóa file ổ cứng
                        }

                        SelectedProfileName = null;
                        SavedProfiles.Remove(toDelete);
                        ProfilePaths.Remove(toDelete);
                        ProfileManager.SaveRegistry(ProfilePaths);
                    }
                },
                canExecute: o => !string.IsNullOrEmpty(SelectedProfileName)
            );

            // Khởi tạo bảng rỗng ban đầu nếu chưa có Profile nào được chọn
            if (string.IsNullOrEmpty(SelectedProfileName))
                CategoryTables.Add(new CategoryTableViewModel(this));
        }

        private void LoadActiveRevitData()
        {
            AvailableLevels.Add("All");
            var levels = new FilteredElementCollector(Doc).OfClass(typeof(Level)).Cast<Level>().Select(l => l.Name).OrderBy(n => n);
            foreach (var lvl in levels) AvailableLevels.Add(lvl);

            var activeCatNames = new FilteredElementCollector(Doc).WhereElementIsNotElementType().Select(e => e.Category).Where(c => c != null && c.CategoryType == CategoryType.Model && c.AllowsBoundParameters).Select(c => c.Name).Distinct().OrderBy(n => n);
            foreach (var c in activeCatNames) AvailableCategories.Add(c);
        }

        // ĐÃ SỬA: CHỤP ẢNH CHÍNH XÁC NHỮNG GÌ ĐANG HIỂN THỊ (Bao gồm Cả Số liệu Các Ô)
        private QtoProfile ExtractDataFromUI(string name)
        {
            var profile = new QtoProfile { ProfileName = name, ProjectData = CurrentProjectInfo };
            foreach (var table in CategoryTables)
            {
                var tProfile = new QtoTableProfile
                {
                    CategoryName = table.SelectedCategory,
                    LevelParameter = table.SelectedLevelParameter,
                    LevelValue = table.SelectedLevelValue,
                    MainParameter = table.MainParameterName,
                    MainHeading = table.MainParameterHeading
                };
                foreach (var col in table.Columns)
                {
                    tProfile.Columns.Add(new QtoColumnProfile { ColumnId = col.ColumnId, ParameterName = col.ParameterName, Heading = col.Heading });
                }

                // Thu thập từng hàng cụ thể và từng ô vuông
                foreach (var row in table.Rows)
                {
                    if (!string.IsNullOrEmpty(row.MainParameterValue))
                    {
                        var rProfile = new QtoRowProfile { MainValue = row.MainParameterValue, Note = row.Note };
                        foreach (var cell in row.Cells)
                        {
                            // Lưu lại giá trị số nguyên chất và giá trị chữ để tiện xuất Excel
                            rProfile.Cells.Add(new QtoCellProfile { StringValue = cell.CellValue, NumericValue = cell.NumericValue, IsNumeric = cell.IsNumeric });
                        }
                        tProfile.Rows.Add(rProfile);
                    }
                }

                profile.Tables.Add(tProfile);
            }
            return profile;
        }

        // ĐÃ SỬA: VẼ LẠI GIAO DIỆN CHÍNH XÁC Y HỆT LÚC LƯU VÀ NẠP TÊN DỰ ÁN
        private void LoadProfileIntoUI(QtoProfile profile)
        {
            if (profile.ProjectData != null) CurrentProjectInfo = profile.ProjectData;

            CategoryTables.Clear(); GlobalColumnIds.Clear(); GrandTotalCells.Clear();

            // Phục hồi Cột Global
            if (profile.Tables.Count > 0)
            {
                foreach (var colProfile in profile.Tables[0].Columns)
                {
                    GlobalColumnIds.Add(colProfile.ColumnId);
                    GrandTotalCells.Add(new GrandTotalCellViewModel(colProfile.ColumnId));
                }
            }

            // Phục hồi các Bảng
            foreach (var tProfile in profile.Tables)
            {
                var newTable = new CategoryTableViewModel(this);
                newTable.SelectedCategory = tProfile.CategoryName;
                newTable.SelectedLevelParameter = tProfile.LevelParameter;
                newTable.SelectedLevelValue = tProfile.LevelValue;

                foreach (var id in GlobalColumnIds) newTable.AddColumn(id);

                newTable.MainParameterName = tProfile.MainParameter;
                newTable.MainParameterHeading = tProfile.MainHeading;

                for (int i = 0; i < tProfile.Columns.Count; i++)
                {
                    var colUI = newTable.Columns.FirstOrDefault(c => c.ColumnId == tProfile.Columns[i].ColumnId);
                    if (colUI != null)
                    {
                        colUI.ParameterName = tProfile.Columns[i].ParameterName;
                        colUI.Heading = tProfile.Columns[i].Heading;
                    }
                }

                // XÓA HÀNG RỖNG MẶC ĐỊNH
                newTable.Rows.Clear();

                // NẠP ĐÚNG CÁC HÀNG CÓ TRONG FILE LƯU
                foreach (var rProfile in tProfile.Rows)
                {
                    var newRow = new RowViewModel(newTable);
                    newRow.MainParameterValue = rProfile.MainValue;
                    newRow.Note = rProfile.Note;
                    newTable.Rows.Add(newRow);
                }

                CategoryTables.Add(newTable);
            }
            RecalculateGrandTotals();
        }

        public void RecalculateGrandTotals()
        {
            foreach (var gtCell in GrandTotalCells)
            {
                var allCellsInCol = CategoryTables.SelectMany(t => t.Rows).SelectMany(r => r.Cells).Where(c => c.Column.ColumnId == gtCell.ColumnId).ToList();
                var validCells = allCellsInCol.Where(c => !string.IsNullOrEmpty(c.CellValue) && c.CellValue != "N/A" && c.CellValue != "0").ToList();

                if (!validCells.Any() || validCells.Any(c => !c.IsNumeric)) { gtCell.DisplayValue = ""; continue; }

                var distinctDataTypes = validCells.Where(c => c.ParamDataType != null).Select(c => c.ParamDataType.TypeId).Distinct().ToList();
                if (distinctDataTypes.Count > 1) { gtCell.DisplayValue = ""; continue; }

                double sum = validCells.Sum(c => c.NumericValue);
                var commonDataType = validCells.FirstOrDefault(c => c.ParamDataType != null)?.ParamDataType;

                if (sum > 0 && commonDataType != null)
                {
                    try { gtCell.DisplayValue = UnitFormatUtils.Format(Doc.GetUnits(), commonDataType, sum, false); }
                    catch { gtCell.DisplayValue = Math.Round(sum, 2).ToString(); }
                }
                else { gtCell.DisplayValue = sum > 0 ? Math.Round(sum, 2).ToString() : ""; }
            }
        }
    }


}