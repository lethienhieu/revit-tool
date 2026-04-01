using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using THBIM; // Namespace chứa RelationshipItem và SyncType

// Alias để tránh trùng tên với Revit API
using WpfContextMenu = System.Windows.Controls.ContextMenu;
using WpfMenuItem = System.Windows.Controls.MenuItem;

namespace THBIM.Views
{
    // --- MODELS UI ---
    public class CategoryItem { public string Name { get; set; } public BuiltInCategory Category { get; set; } }
    public class CategoryRow { public CategoryItem SelectedCategory { get; set; } }
    public class LinkItem { public string Name { get; set; } public bool IsSelected { get; set; } public RevitLinkInstance Instance { get; set; } }

    public partial class StructureSyncWindow : Window
    {
        private readonly StructureSync _logic;
        private ExternalEvent _exEvent;
        private StructureSyncRequestHandler _handler;

        // Cache dữ liệu
        private List<Reference> _currentParentsCacheRefs;
        private SyncType _currentParentType;
        private bool _currentParentsUseLink;

        // Data Collections
        public ObservableCollection<CategoryItem> AllCategoriesList { get; set; }
        public ObservableCollection<CategoryRow> SelectedCategoryRows { get; set; }
        public ObservableCollection<LinkItem> LinkFiles { get; set; }
        public ObservableCollection<RelationshipItem> Relationships { get; set; }

        public StructureSyncWindow(StructureSync logic)
        {
            InitializeComponent();
            _logic = logic;
            this.DataContext = this;

            _handler = new StructureSyncRequestHandler();
            _exEvent = ExternalEvent.Create(_handler);

            InitializeLists();
            LoadLinks();
            LoadSavedData();
        }

        // --- GIAO TIẾP VỚI REVIT (QUA EXTERNAL EVENT) ---
        private void SendRequest(StructureSyncRequestId requestId)
        {
            if (_handler == null || _exEvent == null) return;

            // Nạp dữ liệu vào Handler trước khi Raise
            _handler.Logic = _logic;
            _handler.Request = requestId;
            _handler.Relationships = Relationships.ToList();
            _handler.Links = GetSelectedLinks();

            // Truyền AppExEvent cho Handler (để cửa sổ Report dùng lại cho việc Select)
            _handler.AppExEvent = _exEvent;

            _exEvent.Raise();
        }

        private void LoadSavedData()
        {
            try { var savedItems = _logic.LoadRelationships(); Relationships.Clear(); foreach (var item in savedItems) Relationships.Add(item); } catch { }
        }

        private void SaveData() { SendRequest(StructureSyncRequestId.Save); }

        // --- KHỞI TẠO ---
        private void InitializeLists()
        {
            SelectedCategoryRows = new ObservableCollection<CategoryRow>();
            LinkFiles = new ObservableCollection<LinkItem>();
            Relationships = new ObservableCollection<RelationshipItem>();

            AllCategoriesList = new ObservableCollection<CategoryItem>
            {
                new CategoryItem { Name = "Structural Foundations", Category = BuiltInCategory.OST_StructuralFoundation },
                new CategoryItem { Name = "Structural Framings", Category = BuiltInCategory.OST_StructuralFraming },
                new CategoryItem { Name = "Structural Columns", Category = BuiltInCategory.OST_StructuralColumns },
                new CategoryItem { Name = "Floors", Category = BuiltInCategory.OST_Floors },
                // [UPDATE] Thêm Wall
                new CategoryItem { Name = "Walls", Category = BuiltInCategory.OST_Walls }
            };

            AddCategoryRow();
            icCategories.ItemsSource = SelectedCategoryRows;
            LbLinks.ItemsSource = LinkFiles;
            DgRelationships.ItemsSource = Relationships;
        }

        private void LoadLinks()
        {
            LinkFiles.Clear();
            var links = _logic.GetRevitLinks();
            foreach (var link in links) LinkFiles.Add(new LinkItem { Name = link.Name, IsSelected = false, Instance = link });
        }

        // =========================================================================
        // --- 1. CHỌN PARENT ---
        // =========================================================================

        private void BtnChooseParent_Click(object sender, RoutedEventArgs e)
        {
            var cats = GetSelectedCats();
            if (cats.Count == 0) { MessageBox.Show("Please select at least one Category first!"); return; }

            WpfContextMenu menu = new WpfContextMenu();
            // [UPDATE] Thêm Wall vào Menu
            var types = new List<SyncType> { SyncType.Floor, SyncType.Column, SyncType.PileCap, SyncType.Pile, SyncType.Wall };

            foreach (var t in types)
            {
                WpfMenuItem item = new WpfMenuItem { Header = t.ToString(), FontWeight = FontWeights.Bold };
                item.Click += (s, args) => { RunSelectParentLogic(t); };
                menu.Items.Add(item);
            }

            BtnChooseParent.ContextMenu = menu;
            BtnChooseParent.ContextMenu.IsOpen = true;
        }

        private void RunSelectParentLogic(SyncType type)
        {
            this.Hide();
            try
            {
                var cats = GetSelectedCats();
                bool useLink = ChkUseLinks.IsChecked == true;
                _currentParentsUseLink = useLink;

                var parentRefs = _logic.SelectParents(cats, useLink);
                _currentParentsCacheRefs = parentRefs;
                _currentParentType = type;

                this.Show();

                if (_currentParentsCacheRefs != null && _currentParentsCacheRefs.Count > 0)
                {
                    BtnChooseParent.IsEnabled = false;
                    BtnChooseParent.Content = $"{type} Selected ({_currentParentsCacheRefs.Count}) ✓";
                    BtnChooseChild.IsEnabled = true;
                    BtnChooseChild.Opacity = 1;
                    TxtStatus.Text = $"Parent set as {type}. Now choose Child.";
                }
                else
                {
                    TxtStatus.Text = "Selection canceled.";
                }
            }
            catch (Exception ex)
            {
                this.Show();
                TxtStatus.Text = "Error: " + ex.Message;
            }
        }

        // =========================================================================
        // --- 2. CHỌN CHILD & MATCHING (QUAN TRỌNG: LƯU OFFSET) ---
        // =========================================================================

        private void BtnChooseChild_Click(object sender, RoutedEventArgs e)
        {
            if (_currentParentsCacheRefs == null || _currentParentsCacheRefs.Count == 0) return;

            WpfContextMenu menu = new WpfContextMenu();
            List<SyncType> allowedChildren = new List<SyncType>();

            // [UPDATE] Logic lọc Child hợp lệ
            switch (_currentParentType)
            {
                case SyncType.Floor:
                    allowedChildren.Add(SyncType.DropPanel); break;
                case SyncType.Column:
                    allowedChildren.Add(SyncType.Column); allowedChildren.Add(SyncType.PileCap); allowedChildren.Add(SyncType.DropPanel); allowedChildren.Add(SyncType.Pile); break;
                case SyncType.PileCap:
                    allowedChildren.Add(SyncType.Pile); break;
                case SyncType.Pile:
                    allowedChildren.Add(SyncType.PileCap); break;
                case SyncType.Wall: // [UPDATE] Wall chỉ đi với Wall
                    allowedChildren.Add(SyncType.Wall); break;
                default:
                    allowedChildren.AddRange((SyncType[])Enum.GetValues(typeof(SyncType))); break;
            }

            foreach (var t in allowedChildren)
            {
                WpfMenuItem item = new WpfMenuItem { Header = t.ToString(), FontWeight = FontWeights.Bold };
                item.Click += (s, args) => { RunSelectChildLogic(t); };
                menu.Items.Add(item);
            }

            BtnChooseChild.ContextMenu = menu;
            BtnChooseChild.ContextMenu.IsOpen = true;
        }

        private void RunSelectChildLogic(SyncType childType)
        {
            this.Hide();
            try
            {
                var cats = GetSelectedCats();
                var children = _logic.SelectChildren(cats); // Quét chọn

                this.Show();

                if (children.Count > 0)
                {
                    // [UPDATE] Gọi MatchRelationships với tham số mới nhận OFFSETS
                    _logic.MatchRelationships(_currentParentsCacheRefs, children,
                        _currentParentType, childType,
                        out List<ElementId> matchedParents,
                        out List<ElementId> matchedChildren,
                        out List<string> matchedOffsets); // <--- Output mới

                    if (matchedChildren.Count > 0)
                    {
                        int index = Relationships.Count + 1;
                        string relName = $"SET {index} : \"{_currentParentType}\" > \"{childType}\"";

                        Relationships.Add(new RelationshipItem
                        {
                            Name = relName,
                            ParentTypeStr = _currentParentType.ToString(),
                            ChildTypeStr = childType.ToString(),
                            ParentIsFromLink = _currentParentsUseLink,
                            ChildCount = matchedChildren.Count,
                            ParentIds = matchedParents,
                            ChildIds = matchedChildren,
                            Offsets = matchedOffsets, // [UPDATE] Lưu Offset vào Item
                            IsChecked = true,
                            Status = "New (Pending)"
                        });

                        ResetWorkflowUI();
                        TxtStatus.Text = $"Added {relName}. Press Accept to save.";
                    }
                    else
                    {
                        MessageBox.Show("No intersection found.");
                    }
                }
            }
            catch (Exception ex)
            {
                this.Show();
                TxtStatus.Text = "Error: " + ex.Message;
            }
        }

        // =========================================================================
        // --- BUTTONS ---
        // =========================================================================

        private void BtnTracking_Click(object sender, RoutedEventArgs e)
        {
            if (Relationships.Any(r => r.IsChecked))
            {
                // [UPDATE] Gửi Request REPORT thay vì gọi trực tiếp để tránh Crash
                SendRequest(StructureSyncRequestId.Report);
            }
            else
            {
                MessageBox.Show("Select item first.");
            }
        }

        private void BtnUpdate_Click(object sender, RoutedEventArgs e)
        {
            if (Relationships.Any(r => r.IsChecked)) SendRequest(StructureSyncRequestId.Update);
            else MessageBox.Show("Select item first.");
        }

        private void BtnResetColor_Click(object sender, RoutedEventArgs e)
        {
            if (Relationships.Count > 0) SendRequest(StructureSyncRequestId.ResetColor);
        }

        private void BtnClear_Click(object sender, RoutedEventArgs e)
        {
            if (Relationships.Count > 0 && MessageBox.Show("Clear all unsaved list?", "Confirm", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                Relationships.Clear();
                TxtStatus.Text = "List cleared.";
            }
        }

        private void BtnDeleteRel_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.DataContext is RelationshipItem item)
            {
                Relationships.Remove(item);
                TxtStatus.Text = "Item removed.";
            }
        }

        private void BtnAccept_Click(object sender, RoutedEventArgs e) { SaveData(); this.Close(); }

        // --- HELPERS ---
        private void ResetWorkflowUI()
        {
            BtnChooseChild.IsEnabled = false; BtnChooseChild.Opacity = 0.5;
            BtnChooseParent.IsEnabled = true; BtnChooseParent.Content = "A. Choose Parent";
            _currentParentsCacheRefs = null;
            _currentParentsUseLink = false;
        }

        private void BtnAddCat_Click(object sender, RoutedEventArgs e) => AddCategoryRow();
        private void AddCategoryRow() => SelectedCategoryRows.Add(new CategoryRow { SelectedCategory = AllCategoriesList[0] });
        private void BtnRemoveCat_Click(object sender, RoutedEventArgs e) { if (sender is Button btn && btn.DataContext is CategoryRow row) SelectedCategoryRows.Remove(row); }
        private List<BuiltInCategory> GetSelectedCats() => SelectedCategoryRows.Where(x => x.SelectedCategory != null).Select(x => x.SelectedCategory.Category).Distinct().ToList();
        private List<RevitLinkInstance> GetSelectedLinks() => ChkUseLinks.IsChecked == true ? LinkFiles.Where(x => x.IsSelected).Select(x => x.Instance).ToList() : new List<RevitLinkInstance>();
        private void Header_MouseDown(object sender, MouseButtonEventArgs e) { if (e.LeftButton == MouseButtonState.Pressed) DragMove(); }
        private void BtnClose_Click(object sender, RoutedEventArgs e) => Close();
        private void ChkUseLinks_Changed(object sender, RoutedEventArgs e) { }
    }
}
