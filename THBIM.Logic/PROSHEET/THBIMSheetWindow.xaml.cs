using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using DB = Autodesk.Revit.DB;

namespace THBIM
{
    public partial class THBIMSheetWindow : Window
    {
        private DB.Document _doc;
        private ExternalEvent _exEvent;
        private RequestHandler _handler;
        public ObservableCollection<SheetItem> CombineOrderList { get; set; } = new ObservableCollection<SheetItem>();

        private Dictionary<string, string> _profilePaths = new Dictionary<string, string>();
        private bool _suppressProfileSelection = false;

        private static string ProfileFolder
        {
            get
            {
                string folder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "THBIM", "ProSheet", "Profiles");
                if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);
                return folder;
            }
        }

        private static string LastProfilePath =>
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "THBIM", "ProSheet", "last_profile.txt");

        private static void SaveLastProfileName(string name)
        {
            try { File.WriteAllText(LastProfilePath, name); } catch { }
        }

        private static string LoadLastProfileName()
        {
            try { return File.Exists(LastProfilePath) ? File.ReadAllText(LastProfilePath).Trim() : null; } catch { return null; }
        }

        public THBIMSheetWindow(DB.Document doc, ExternalEvent exEvent, RequestHandler handler)
        {

            InitializeComponent();
            _doc = doc;
            _exEvent = exEvent;
            _handler = handler;

            // =======================================================
            // SỬA LỖI RESET BẢNG: Chỉ Refresh UI khi tạo/xóa Set, KHÔNG Refresh khi Export
            // =======================================================
            _handler.OnRequestCompleted = () =>
            {
                this.Dispatcher.Invoke(() =>
                {
                    // Kiểm tra: Nếu KHÔNG PHẢI là lệnh Export thì mới quét và làm mới lại bảng
                    if (_handler.Req.Type != RequestType.Export)
                    {
                        RefreshSelectionUI();
                    }
                    else
                    {
                        // Nếu bạn có viết code làm mờ nút Export lúc đang chạy (BtnExport.IsEnabled = false),
                        // thì bạn có thể bật lại nó ở ngay đây:
                        BtnNext.IsEnabled = true;
                    }
                });
            };
            _handler.OnProgressUpdated = (completed, total) =>
            {
                // Code này đã chạy ngầm trong Dispatcher nên update trực tiếp luôn
                PbOverall.Maximum = total;
                PbOverall.Value = completed;
                TxtOverallProgress.Text = $"Completed {completed} of {total} tasks";
            };

            this.DataContext = this;

            // Load danh sách profile từ thư mục
            LoadProfileList();

            // Khởi tạo dữ liệu cho các tab
            //InitSelectionTab();
            //InitFormatTab();
            this.ContentRendered += async (s, e) => { await InitSelectionTabAsync(); };
        }
        // 1. Logic làm mờ/sáng nút dựa theo Tab đang được chọn
        private void Tab_Checked(object sender, RoutedEventArgs e)
        {
            if (!this.IsLoaded) return;
            if (BtnBack == null || BtnNext == null) return;

            if (TabSelection.IsChecked == true)
            {
                BtnBack.IsEnabled = false;
                BtnNext.IsEnabled = true;
            }
            else if (TabFormat.IsChecked == true)
            {
                BtnBack.IsEnabled = true;
                BtnNext.IsEnabled = true;

                // GỌI HÀM NÀY ĐỂ CẬP NHẬT DANH SÁCH GỘP FILE
                UpdateCombineList();
            }
            else if (TabCreate.IsChecked == true)
            {
                BtnBack.IsEnabled = true;
                BtnNext.IsEnabled = true;
                BtnNext.Content = "Create"; // Đổi chữ nút Next thành Create để chuẩn bị xuất file

                // GỌI HÀM NÀY ĐỂ TỰ ĐỘNG NẠP DANH SÁCH BẢN VẼ
                UpdateCreateList();
            }
        }

        // 2. Logic điều hướng khi bấm Next
        // Sự kiện khi bấm nút Next / Create ở góc dưới cùng bên phải
        private void BtnNext_Click(object sender, RoutedEventArgs e)
        {
            if (TabSelection.IsChecked == true)
            {
                if (!_masterList.Any(x => x.IsSelected))
                {
                    MessageBox.Show("Please select at least one Sheet or View to continue!", "Notice", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return; // Dừng luôn, không cho chạy lệnh chuyển tab bên dưới
                }
                TabFormat.IsChecked = true; // Chuyển sang Format
            }
            else if (TabFormat.IsChecked == true)
            {
                TabCreate.IsChecked = true; // Chuyển sang Create
            }
            else if (TabCreate.IsChecked == true)
            {
                // NẾU ĐANG Ở TAB CREATE MÀ BẤM NÚT NÀY -> BẮT ĐẦU XUẤT FILE
                StartExportProcess();
            }
        }

        // 3. Logic điều hướng khi bấm Back
        private void BtnBack_Click(object sender, RoutedEventArgs e)
        {
            if (TabFormat.IsChecked == true)
            {
                // Lùi về Tab Selection
                TabSelection.IsChecked = true;
            }
            else if (TabCreate.IsChecked == true)
            {
                // Lùi về Tab Format
                TabFormat.IsChecked = true;
            }
        }
        private void BtnEditOrder_Click(object sender, RoutedEventArgs e)
        {
            if (CombineOrderList == null) CombineOrderList = new System.Collections.ObjectModel.ObservableCollection<SheetItem>();

            // Gọi cửa sổ mới và ném danh sách bản vẽ vào cho nó
            var dialog = new CombineOrderDialog(CombineOrderList);

            // Chờ người dùng bấm nút OK ở cửa sổ đó
            if (dialog.ShowDialog() == true)
            {
                // Thay thế danh sách cũ bằng danh sách đã được sắp xếp mới
                CombineOrderList.Clear();
                foreach (var item in dialog.OrderedList)
                {
                    CombineOrderList.Add(item);
                }
            }
        }

        private void TabControl_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {

        }

        // Sự kiện khi người dùng chọn Khổ Giấy (A1, A0...)
        private void MenuPaperSize_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as System.Windows.Controls.MenuItem;
            if (menuItem == null) return;

            string selectedSize = menuItem.Header.ToString();

            if (DgCreate.SelectedItems.Count == 0)
            {
                MessageBox.Show("No sheets selected for paper size application.", "Warning", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            foreach (var item in DgCreate.SelectedItems)
            {
                // Dùng ExportTaskItem thay vì SheetItem
                if (item is ExportTaskItem taskItem)
                {
                    taskItem.PaperSize = selectedSize;
                }
            }
        }

        // Sự kiện khi người dùng chọn Chiều Giấy (Landscape / Portrait)
        private void MenuOrientation_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as System.Windows.Controls.MenuItem;
            if (menuItem == null) return;

            // Đặt tên biến là selectedOrientation (tránh nhầm lẫn với selectedSize)
            string selectedOrientation = menuItem.Header.ToString();

            if (DgCreate.SelectedItems.Count == 0) return;

            foreach (var item in DgCreate.SelectedItems)
            {
                if (item is ExportTaskItem taskItem)
                {
                    taskItem.Orientation = selectedOrientation;
                }
            }
        }

        // ==========================================================
        // PROFILE MANAGEMENT (giống SheetLink)
        // ==========================================================

        private void LoadProfileList()
        {
            _suppressProfileSelection = true;
            CboProfiles.Items.Clear();

            // Thêm placeholder "Please Select"
            var placeholder = new System.Windows.Controls.ComboBoxItem { Content = "Please Select" };
            CboProfiles.Items.Add(placeholder);
            CboProfiles.SelectedItem = placeholder;

            _profilePaths.Clear();

            try
            {
                foreach (var file in Directory.GetFiles(ProfileFolder, "*.txt").OrderBy(f => f))
                {
                    string name = Path.GetFileNameWithoutExtension(file);
                    _profilePaths[name] = file;
                    CboProfiles.Items.Add(new System.Windows.Controls.ComboBoxItem { Content = name });
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"LoadProfileList error: {ex.Message}");
            }
            finally
            {
                _suppressProfileSelection = false;
            }

        }

        private void CboProfiles_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (_suppressProfileSelection) return;
            if (CboProfiles.SelectedItem is not System.Windows.Controls.ComboBoxItem item) return;
            var name = item.Content?.ToString();
            if (string.IsNullOrEmpty(name) || name == "Please Select") return;

            if (!_profilePaths.TryGetValue(name, out string path)) return;

            try
            {
                string content = File.ReadAllText(path);
                var data = ParseTxtProfile(content);
                if (data != null) ApplyTxtProfileToUI(data);

                // Ghi nhớ profile vừa chọn
                SaveLastProfileName(name);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to load profile '{name}': {ex.Message}",
                    "ProSheet", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // 💾 Save button → ContextMenu (Save / Save As...) giống SheetLink
        private void BtnSaveProfile_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not System.Windows.Controls.Button btn)
            {
                SaveProfile();
                return;
            }

            var menu = new System.Windows.Controls.ContextMenu
            {
                PlacementTarget = btn,
                Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom
            };
            var miSave = new System.Windows.Controls.MenuItem { Header = "Save" };
            miSave.Click += (_, _) => SaveProfile();
            var miSaveAs = new System.Windows.Controls.MenuItem { Header = "Save As..." };
            miSaveAs.Click += (_, _) => SaveProfileAs();
            menu.Items.Add(miSave);
            menu.Items.Add(miSaveAs);
            menu.IsOpen = true;
        }

        private void SaveProfile()
        {
            var name = GetCurrentProfileName();
            if (string.IsNullOrWhiteSpace(name))
                name = ShowInputDialog("Enter profile name:", "Save Profile", "New Profile");
            if (string.IsNullOrWhiteSpace(name)) return;

            try
            {
                string txt = SnapshotToTxt(name);
                string path = Path.Combine(ProfileFolder, GetSafeFileName(name) + ".txt");
                File.WriteAllText(path, txt);

                LoadProfileList();
                SelectProfile(name);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "ProSheet",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SaveProfileAs()
        {
            var defaultName = GetCurrentProfileName();
            if (string.IsNullOrWhiteSpace(defaultName))
                defaultName = $"ProSheet_Profile_{DateTime.Now:yyyyMMdd_HHmm}";

            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Title = "Save Profile As",
                Filter = "Text Files|*.txt|All Files|*.*",
                FileName = defaultName + ".txt",
                AddExtension = true,
                DefaultExt = ".txt",
                OverwritePrompt = true
            };
            if (dlg.ShowDialog() != true) return;

            try
            {
                var fileName = Path.GetFileNameWithoutExtension(dlg.FileName);
                string txt = SnapshotToTxt(fileName);

                // Export to external file
                File.WriteAllText(dlg.FileName, txt);

                // Also save to internal profiles folder
                string internalPath = Path.Combine(ProfileFolder, GetSafeFileName(fileName) + ".txt");
                File.WriteAllText(internalPath, txt);

                LoadProfileList();
                SelectProfile(fileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "ProSheet",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // 📥 Import Profile (giống SheetLink BtnImportProfile_Click)
        private void BtnImportProfile_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Title = "Import Profile",
                Filter = "Text Files|*.txt|All Files|*.*",
                CheckFileExists = true
            };
            if (dlg.ShowDialog() != true) return;

            try
            {
                string content = File.ReadAllText(dlg.FileName);
                var data = ParseTxtProfile(content);
                if (data == null || !data.ContainsKey("Name"))
                {
                    MessageBox.Show("The profile file is invalid.", "ProSheet",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                string profileName = data.TryGetValue("Name", out var n) ? n : Path.GetFileNameWithoutExtension(dlg.FileName);
                if (string.IsNullOrWhiteSpace(profileName))
                    profileName = Path.GetFileNameWithoutExtension(dlg.FileName);

                // Kiểm tra trùng tên
                if (_profilePaths.ContainsKey(profileName))
                {
                    var overwrite = MessageBox.Show(
                        $"Profile '{profileName}' already exists. Overwrite?",
                        "ProSheet",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Question);
                    if (overwrite != MessageBoxResult.Yes) return;
                }

                // Save to internal profiles folder
                string path = Path.Combine(ProfileFolder, GetSafeFileName(profileName) + ".txt");
                File.WriteAllText(path, content);

                LoadProfileList();
                SelectProfile(profileName);
                ApplyTxtProfileToUI(data);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "ProSheet",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // 🗑 Delete Profile (giống SheetLink)
        private void BtnDeleteProfile_Click(object sender, RoutedEventArgs e)
        {
            if (CboProfiles.SelectedItem is not System.Windows.Controls.ComboBoxItem item) return;
            var name = item.Content?.ToString();
            if (string.IsNullOrEmpty(name) || name == "Please Select") return;

            var res = MessageBox.Show($"Delete profile \"{name}\"?",
                "ProSheet", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (res != MessageBoxResult.Yes) return;

            try
            {
                if (_profilePaths.TryGetValue(name, out string path) && File.Exists(path))
                    File.Delete(path);
            }
            catch { }

            LoadProfileList();
        }

        // ── Helpers ──────────────────────────────────────────────────

        private string GetCurrentProfileName()
        {
            if (CboProfiles.SelectedItem is System.Windows.Controls.ComboBoxItem item)
            {
                var fromCombo = item.Content?.ToString();
                if (!string.IsNullOrWhiteSpace(fromCombo) &&
                    !string.Equals(fromCombo, "Please Select", StringComparison.OrdinalIgnoreCase))
                    return fromCombo;
            }
            return string.Empty;
        }

        private void SelectProfile(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return;
            _suppressProfileSelection = true;
            foreach (System.Windows.Controls.ComboBoxItem item in CboProfiles.Items)
            {
                if (string.Equals(item.Content?.ToString(), name, StringComparison.OrdinalIgnoreCase))
                {
                    CboProfiles.SelectedItem = item;
                    break;
                }
            }
            _suppressProfileSelection = false;
        }

        private static string GetSafeFileName(string name)
        {
            return new string(name.Select(c =>
                Path.GetInvalidFileNameChars().Contains(c) ? '_' : c).ToArray());
        }

        private static string ShowInputDialog(string prompt, string title, string defaultValue = "")
        {
            var dialog = new Window
            {
                Title = title,
                Width = 400, Height = 160,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                WindowStyle = WindowStyle.ToolWindow,
                ResizeMode = ResizeMode.NoResize
            };

            var stack = new System.Windows.Controls.StackPanel { Margin = new Thickness(15) };
            stack.Children.Add(new System.Windows.Controls.TextBlock { Text = prompt, Margin = new Thickness(0, 0, 0, 8) });

            var textBox = new System.Windows.Controls.TextBox
            {
                Text = defaultValue,
                Height = 28,
                VerticalContentAlignment = VerticalAlignment.Center,
                Padding = new Thickness(5, 0, 5, 0)
            };
            stack.Children.Add(textBox);

            var btnPanel = new System.Windows.Controls.StackPanel
            {
                Orientation = System.Windows.Controls.Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 12, 0, 0)
            };

            var btnOk = new System.Windows.Controls.Button
            {
                Content = "OK", Width = 80, Height = 28, Margin = new Thickness(0, 0, 8, 0),
                Background = new System.Windows.Media.SolidColorBrush(
                    (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#2980B9")),
                Foreground = System.Windows.Media.Brushes.White,
                BorderThickness = new Thickness(0), Cursor = System.Windows.Input.Cursors.Hand
            };
            btnOk.Click += (_, _) => { dialog.DialogResult = true; };

            var btnCancel = new System.Windows.Controls.Button
            {
                Content = "Cancel", Width = 80, Height = 28, Cursor = System.Windows.Input.Cursors.Hand
            };
            btnCancel.Click += (_, _) => { dialog.DialogResult = false; };

            btnPanel.Children.Add(btnOk);
            btnPanel.Children.Add(btnCancel);
            stack.Children.Add(btnPanel);
            dialog.Content = stack;

            textBox.SelectAll();
            textBox.Focus();

            return dialog.ShowDialog() == true && !string.IsNullOrWhiteSpace(textBox.Text)
                ? textBox.Text.Trim() : null;
        }

        // ==========================================================
        // SNAPSHOT → TXT (lưu toàn bộ trạng thái UI)
        // ==========================================================

        private string SnapshotToTxt(string name)
        {
            var lines = new List<string>();
            var now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            // ── Metadata ──
            lines.Add($"Name={name}");
            lines.Add($"CreatedAt={now}");
            lines.Add($"UpdatedAt={now}");

            // ── Selection Tab ──
            lines.Add($"Mode={((RbSheets?.IsChecked == true) ? "Sheets" : "Views")}");

            // Lưu naming rule format string nếu có
            if (_savedNamingRules != null && _savedNamingRules.Count > 0)
            {
                var ruleStr = string.Join("|", _savedNamingRules.Select(r =>
                    $"{r.ParameterName}~{r.Prefix}~{r.Suffix}~{r.Separator}"));
                lines.Add($"NamingRules={ruleStr}");
            }

            // ── Format Tab - PDF ──
            lines.Add($"ExportPDF={ChkExportPDF.IsChecked == true}");
            lines.Add($"ExportDWG={ChkExportDWG.IsChecked == true}");
            lines.Add($"IsCenter={RbCenter.IsChecked == true}");
            lines.Add($"OffsetX={TxtOffsetX.Text}");
            lines.Add($"OffsetY={TxtOffsetY.Text}");
            lines.Add($"IsFitToPage={RbFitToPage.IsChecked == true}");
            lines.Add($"ZoomPercentage={TxtZoomPercentage.Text}");
            lines.Add($"RasterQuality={CboRasterQuality.SelectedIndex}");
            lines.Add($"Colors={CboColors.SelectedIndex}");
            lines.Add($"HideRefPlanes={ChkHideRefPlanes.IsChecked == true}");
            lines.Add($"HideUnrefTags={ChkHideUnrefTags.IsChecked == true}");
            lines.Add($"HideCrop={ChkHideCrop.IsChecked == true}");
            lines.Add($"HideScopeBox={ChkHideScopeBox.IsChecked == true}");
            lines.Add($"IsSeparateFile={RbSeparate.IsChecked == true}");
            lines.Add($"CombinedName={TxtCombinedName.Text}");

            // ── Format Tab - DWG ──
            string dwgSetup = "";
            if (CboDwgExportSetup.SelectedItem is ExportSetupItem setupItem)
                dwgSetup = setupItem.Name;
            lines.Add($"DWGSettingName={dwgSetup}");
            lines.Add($"DWG_MergedViews={ChkExportViewsOnSheets.IsChecked == true}");
            lines.Add($"DWG_BindImages={ChkBindImages.IsChecked == true}");
            lines.Add($"CleanPcp={ChkCleanPcp.IsChecked == true}");

            // ── Create Tab ──
            lines.Add($"ExportFolder={TxtExportFolder.Text}");
            lines.Add($"SplitByFormat={RbSaveSplitFormat.IsChecked == true}");

            return string.Join(Environment.NewLine, lines);
        }

        // ==========================================================
        // APPLY TXT PROFILE → UI
        // ==========================================================

        private void ApplyTxtProfileToUI(Dictionary<string, string> d)
        {
            if (d == null) return;

            // ── Selection Tab ──
            // Không khôi phục tick chọn — chỉ khôi phục naming rules + custom file names

            // Khôi phục naming rules
            var rulesStr = Get(d, "NamingRules");
            if (!string.IsNullOrEmpty(rulesStr))
            {
                _savedNamingRules = new ObservableCollection<NameRuleItem>();
                foreach (var part in rulesStr.Split('|'))
                {
                    var segs = part.Split('~');
                    if (segs.Length >= 4)
                    {
                        _savedNamingRules.Add(new NameRuleItem
                        {
                            ParameterName = segs[0],
                            Prefix = segs[1],
                            Suffix = segs[2],
                            Separator = segs[3]
                        });
                    }
                }

                // Re-evaluate custom file name cho TOÀN BỘ masterList dựa trên naming rules
                string rawFormat = BuildFormatStringFromRules(_savedNamingRules);
                if (!string.IsNullOrEmpty(rawFormat))
                {
                    foreach (var item in _masterList)
                    {
                        var elem = _doc.GetElement(item.Id);
                        if (elem != null)
                            item.CustomFileName = EvaluateFileNameFormat(elem, rawFormat);
                    }
                }

                DgSheets?.Items.Refresh();
            }

            // ── Format Tab - PDF ──
            ChkExportPDF.IsChecked = GetBool(d, "ExportPDF");
            ChkExportDWG.IsChecked = GetBool(d, "ExportDWG");

            RbCenter.IsChecked = GetBool(d, "IsCenter", true);
            RbOffset.IsChecked = !GetBool(d, "IsCenter", true);

            var ox = Get(d, "OffsetX");
            var oy = Get(d, "OffsetY");
            if (!string.IsNullOrEmpty(ox)) TxtOffsetX.Text = ox;
            if (!string.IsNullOrEmpty(oy)) TxtOffsetY.Text = oy;

            RbFitToPage.IsChecked = GetBool(d, "IsFitToPage", true);
            RbZoom.IsChecked = !GetBool(d, "IsFitToPage", true);

            var zoom = Get(d, "ZoomPercentage");
            if (!string.IsNullOrEmpty(zoom)) TxtZoomPercentage.Text = zoom;

            if (int.TryParse(Get(d, "RasterQuality"), out int rq) && rq >= 0 && rq < CboRasterQuality.Items.Count)
                CboRasterQuality.SelectedIndex = rq;

            if (int.TryParse(Get(d, "Colors"), out int ci) && ci >= 0 && ci < CboColors.Items.Count)
                CboColors.SelectedIndex = ci;

            ChkHideRefPlanes.IsChecked = GetBool(d, "HideRefPlanes");
            ChkHideUnrefTags.IsChecked = GetBool(d, "HideUnrefTags");
            ChkHideCrop.IsChecked = GetBool(d, "HideCrop");
            ChkHideScopeBox.IsChecked = GetBool(d, "HideScopeBox");

            bool isSep = GetBool(d, "IsSeparateFile", true);
            RbSeparate.IsChecked = isSep;
            RbCombine.IsChecked = !isSep;

            var cname2 = Get(d, "CombinedName");
            if (!string.IsNullOrEmpty(cname2)) TxtCombinedName.Text = cname2;

            // ── Format Tab - DWG ──
            var dwgSetupName = Get(d, "DWGSettingName");
            if (!string.IsNullOrEmpty(dwgSetupName) && CboDwgExportSetup.ItemsSource != null)
            {
                foreach (var item in CboDwgExportSetup.ItemsSource)
                {
                    if (item is ExportSetupItem setup && setup.Name == dwgSetupName)
                    { CboDwgExportSetup.SelectedItem = item; break; }
                }
            }

            ChkExportViewsOnSheets.IsChecked = GetBool(d, "DWG_MergedViews");
            ChkBindImages.IsChecked = GetBool(d, "DWG_BindImages");
            ChkCleanPcp.IsChecked = GetBool(d, "CleanPcp");

            // ── Create Tab ──
            var folder = Get(d, "ExportFolder");
            if (!string.IsNullOrEmpty(folder)) TxtExportFolder.Text = folder;

            bool split = GetBool(d, "SplitByFormat");
            RbSaveSplitFormat.IsChecked = split;
            RbSaveSameFolder.IsChecked = !split;
        }

        // ==========================================================
        // TXT PARSE HELPERS (giống SheetLink)
        // ==========================================================

        private static Dictionary<string, string> ParseTxtProfile(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return null;
            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var line in text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var idx = line.IndexOf('=');
                if (idx <= 0) continue;
                dict[line[..idx].Trim()] = line[(idx + 1)..];
            }
            return dict;
        }

        private static string Get(Dictionary<string, string> d, string key, string def = "")
            => d.TryGetValue(key, out var v) && !string.IsNullOrEmpty(v) ? v : def;

        private static bool GetBool(Dictionary<string, string> d, string key, bool def = false)
            => d.TryGetValue(key, out var v) && bool.TryParse(v, out var b) ? b : def;

        private static List<string> GetList(Dictionary<string, string> d, string key)
        {
            if (!d.TryGetValue(key, out var v) || string.IsNullOrWhiteSpace(v))
                return new List<string>();
            return v.Split('|').Where(s => !string.IsNullOrEmpty(s)).ToList();
        }

        /// <summary>
        /// Build format string từ naming rules (giống CustomNameDialog.UpdatePreview)
        /// Ví dụ: rules [{Sheet Number, "", "", "-"}, {Sheet Name, "", "", ""}]
        /// → "&lt;Sheet Number&gt;-&lt;Sheet Name&gt;"
        /// </summary>
        private static string BuildFormatStringFromRules(IList<NameRuleItem> rules)
        {
            if (rules == null || rules.Count == 0) return "";
            var parts = new List<string>();
            for (int i = 0; i < rules.Count; i++)
            {
                var r = rules[i];
                string part = $"{r.Prefix}<{r.ParameterName}>{r.Suffix}";
                if (i < rules.Count - 1)
                    part += r.Separator;
                parts.Add(part);
            }
            return string.Join("", parts);
        }
    }
}
