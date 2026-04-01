using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Media.Animation;
using System.Windows.Threading;
using THBIM.Controls;
using THBIM.Helpers;
using THBIM.Models;
using THBIM.Services;

namespace THBIM
{
    public partial class SheetLinkWindow : Window
    {
        private readonly Services.ProfileManager _profileMgr = Services.ProfileManager.Instance;
        private readonly SheetLinkState _state = SheetLinkState.Instance;
        private ProgressAnimator _progress;
        private BusyOverlay _overlay;
        private DispatcherTimer _previewTimer;
        private PreviewPayload _previewPayload = PreviewPayload.Empty;
        private string _previewPreferredTag = "Model";
        private readonly Dictionary<ComboBox, (Brush Bg, Brush Border, Effect Effect)> _comboDefaultVisuals = new();
        private readonly Dictionary<TextBox, (Brush Bg, Brush Border)> _comboEditableDefaults = new();
        private readonly Dictionary<ToggleButton, (Brush Bg, Brush Border)> _comboToggleDefaults = new();
        private readonly HashSet<Button> _hoverButtons = new();
        private readonly Dictionary<Button, (Brush Bg, Brush Border, Effect Effect, Transform Transform)> _buttonDefaultVisuals = new();

        private sealed class PreviewPayload
        {
            public static readonly PreviewPayload Empty = new(
                "",
                "",
                new System.Collections.Generic.List<string>(),
                new System.Collections.Generic.List<string>(),
                new System.Collections.Generic.List<PreviewRowData>(),
                new System.Collections.Generic.Dictionary<string, ParamKind>(StringComparer.OrdinalIgnoreCase));

            public PreviewPayload(
                string tag,
                string title,
                System.Collections.Generic.IReadOnlyList<string> targets,
                System.Collections.Generic.IReadOnlyList<string> parameters,
                System.Collections.Generic.IReadOnlyList<PreviewRowData> rows,
                System.Collections.Generic.IReadOnlyDictionary<string, ParamKind> paramKindMap = null)
            {
                Tag = tag ?? string.Empty;
                Title = title ?? string.Empty;
                Targets = targets ?? new System.Collections.Generic.List<string>();
                Parameters = parameters ?? new System.Collections.Generic.List<string>();
                Rows = rows ?? new System.Collections.Generic.List<PreviewRowData>();
                ParamKindMap = paramKindMap ?? new System.Collections.Generic.Dictionary<string, ParamKind>(StringComparer.OrdinalIgnoreCase);
            }

            public string Tag { get; }
            public string Title { get; }
            public System.Collections.Generic.IReadOnlyList<string> Targets { get; }
            public System.Collections.Generic.IReadOnlyList<string> Parameters { get; }
            public System.Collections.Generic.IReadOnlyList<PreviewRowData> Rows { get; }
            public System.Collections.Generic.IReadOnlyDictionary<string, ParamKind> ParamKindMap { get; }
            public bool HasData => Parameters.Count > 0;
        }

        public SheetLinkWindow()
        {
            InitializeComponent();
            Loaded += OnLoaded;
            Closing += OnClosing;
            RegisterShortcuts();
            AddHandler(ComboBox.PreviewMouseLeftButtonDownEvent, new MouseButtonEventHandler(ComboBox_PreviewMouseLeftButtonDown), true);
            AddHandler(TextBox.PreviewMouseMoveEvent, new MouseEventHandler(ComboBoxEditableText_PreviewMouseMove), true);
            AddHandler(ComboBox.MouseEnterEvent, new MouseEventHandler(ComboBox_MouseEnter), true);
            AddHandler(ComboBox.MouseLeaveEvent, new MouseEventHandler(ComboBox_MouseLeave), true);
        }

        // ══════════════════════════════════════════════════════════════════
        // INIT
        // ══════════════════════════════════════════════════════════════════

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            _progress = new ProgressAnimator(ProgressFill, TbProgress);

            _overlay = new BusyOverlay();
            if (Content is Grid rootGrid)
                rootGrid.Children.Add(_overlay);

            LoadProfileList();

            // Auto-restore last-used profile from Revit document storage
            try
            {
                var doc = Services.RevitDocumentCache.Current;
                if (doc != null)
                {
                    var savedProfile = Services.RevitStorageService.LoadFromDocument(doc);
                    if (savedProfile != null && !string.IsNullOrWhiteSpace(savedProfile.Name))
                    {
                        _profileMgr.ApplyToState(savedProfile);
                        SelectProfile(savedProfile.Name);
                        ApplyProfileToViews(savedProfile);
                    }
                }
            }
            catch { /* ignore — DB load is best-effort */ }

            ViewModelCategories.Visibility = Visibility.Visible;
            RefreshPreviewTabState();
            ApplyComboBoxCursorMode(this);
            ApplyButtonHoverMode(this);

            _previewTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(450)
            };
            _previewTimer.Tick += (_, _) => RefreshPreviewTabState();
            _previewTimer.Start();

            Opacity = 0;
            BeginAnimation(OpacityProperty,
                new DoubleAnimation(0, 1,
                    new Duration(TimeSpan.FromMilliseconds(200))));
        }

        private void OnClosing(object sender,
                                System.ComponentModel.CancelEventArgs e)
        {
            _previewTimer?.Stop();

            if (_overlay?.Visibility == Visibility.Visible)
            {
                var res = MessageBox.Show(
                    "Processing is still running — are you sure you want to close?",
                    "SheetLink", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (res == MessageBoxResult.No) e.Cancel = true;
            }
        }

        // ══════════════════════════════════════════════════════════════════
        // TITLE BAR
        // ══════════════════════════════════════════════════════════════════

        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2) ToggleMaximize();
            else DragMove();
        }

        private void BtnMinimize_Click(object sender, RoutedEventArgs e)
            => WindowState = WindowState.Minimized;

        private void BtnMaximize_Click(object sender, RoutedEventArgs e)
            => ToggleMaximize();

        private void BtnClose_Click(object sender, RoutedEventArgs e)
            => Close();

        private void ToggleMaximize()
        {
            if (WindowState == WindowState.Maximized)
            {
                WindowState = WindowState.Normal;
                BtnMaxRestore.Content = "□";
            }
            else
            {
                WindowState = WindowState.Maximized;
                BtnMaxRestore.Content = "❐";
            }
        }

        // ══════════════════════════════════════════════════════════════════
        // TAB SWITCHING
        // ══════════════════════════════════════════════════════════════════

        public void Tab_Click(object sender, RoutedEventArgs e)
        {
            RefreshPreviewTabState();

            var tag = (sender as RadioButton)?.Tag?.ToString();
            if (string.Equals(tag, "Preview", StringComparison.OrdinalIgnoreCase) && !TabPreview.IsEnabled)
            {
                SetProgressWarning("Preview is only available after selecting both an object and its parameters.");
                TabModel.IsChecked = true;
                tag = "Model";
            }
            if (!string.IsNullOrWhiteSpace(tag) && !string.Equals(tag, "Preview", StringComparison.OrdinalIgnoreCase))
                _previewPreferredTag = tag;

            ViewModelCategories.Visibility = Visibility.Collapsed;
            ViewAnnotation.Visibility = Visibility.Collapsed;
            ViewElements.Visibility = Visibility.Collapsed;
            ViewSchedules.Visibility = Visibility.Collapsed;
            ViewSpatial.Visibility = Visibility.Collapsed;
            ViewPreview.Visibility = Visibility.Collapsed;

            switch (tag)
            {
                case "Model": ViewModelCategories.Visibility = Visibility.Visible; break;
                case "Annotation": ViewAnnotation.Visibility = Visibility.Visible; break;
                case "Elements": ViewElements.Visibility = Visibility.Visible; break;
                case "Schedules": ViewSchedules.Visibility = Visibility.Visible; break;
                case "Spatial": ViewSpatial.Visibility = Visibility.Visible; break;
                case "Preview":
                    ViewPreview.Visibility = Visibility.Visible;
                    ViewPreview.LoadPreviewData(_previewPayload.Title, _previewPayload.Targets, _previewPayload.Parameters, _previewPayload.Rows, _previewPayload.ParamKindMap);
                    break;
            }

            ApplyComboBoxCursorMode(this);
            ApplyButtonHoverMode(this);
            _progress?.Reset();
        }

        private void SwitchToTab(string tag)
        {
            var map = new System.Collections.Generic.Dictionary<string, RadioButton>
            {
                { "Model",      TabModel      },
                { "Annotation", TabAnnotation },
                { "Elements",   TabElements   },
                { "Schedules",  TabSchedules  },
                { "Spatial",    TabSpatial    },
                { "Preview",    TabPreview    }
            };
            if (map.TryGetValue(tag, out var tab))
            {
                if (ReferenceEquals(tab, TabPreview))
                {
                    RefreshPreviewTabState();
                    if (!TabPreview.IsEnabled) return;
                }

                tab.IsChecked = true;
                Tab_Click(tab, new RoutedEventArgs());
            }
        }

        public bool TryOpenPreview()
        {
            var activeTag = GetActiveTabTag();
            if (!string.Equals(activeTag, "Preview", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(activeTag))
                _previewPreferredTag = activeTag;

            RefreshPreviewTabState();
            if (!TabPreview.IsEnabled)
            {
                SetProgressWarning("Preview is not ready yet. Please select an object and its parameters first.");
                return false;
            }

            TabPreview.IsChecked = true;
            Tab_Click(TabPreview, new RoutedEventArgs());
            return true;
        }

        private void RefreshPreviewTabState()
        {
            _previewPayload = BuildPreviewPayload();

            TabPreview.IsEnabled = _previewPayload.HasData;
            TabPreview.Opacity = _previewPayload.HasData ? 1.0 : 0.5;

            if (!_previewPayload.HasData && TabPreview.IsChecked == true)
            {
                TabModel.IsChecked = true;
                Tab_Click(TabModel, new RoutedEventArgs());
            }

            ViewPreview?.LoadPreviewData(_previewPayload.Title, _previewPayload.Targets, _previewPayload.Parameters, _previewPayload.Rows, _previewPayload.ParamKindMap);
        }

        private PreviewPayload BuildPreviewPayload()
        {
            var all = new System.Collections.Generic.List<PreviewPayload>
            {
                CreatePayload("Model", "Model Categories", ViewModelCategories?.GetPreviewTargets(), ViewModelCategories?.GetPreviewParameters(), ViewModelCategories?.GetPreviewRows(), ViewModelCategories?.GetSelectedParameterItems()),
                CreatePayload("Annotation", "Annotation Categories", ViewAnnotation?.GetPreviewTargets(), ViewAnnotation?.GetPreviewParameters(), ViewAnnotation?.GetPreviewRows(), ViewAnnotation?.GetSelectedParameterItems()),
                CreatePayload("Elements", "Elements", ViewElements?.GetPreviewTargets(), ViewElements?.GetPreviewParameters(), ViewElements?.GetPreviewRows(), ViewElements?.GetSelectedParameterItems()),
                CreatePayload("Schedules", "Schedules", ViewSchedules?.GetPreviewTargets(), ViewSchedules?.GetPreviewParameters(), ViewSchedules?.GetPreviewRows(), ViewSchedules?.GetSelectedParameterItems()),
                CreatePayload("Spatial", ViewSpatial?.GetPreviewSourceLabel() ?? "Spatial", ViewSpatial?.GetPreviewTargets(), ViewSpatial?.GetPreviewParameters(), ViewSpatial?.GetPreviewRows(), ViewSpatial?.GetSelectedParameterItems())
            };

            var activeTag = ResolvePreviewSourceTag();
            var active = all.FirstOrDefault(x => x.HasData && string.Equals(x.Tag, activeTag, StringComparison.OrdinalIgnoreCase));
            return active ?? PreviewPayload.Empty;
        }

        private static PreviewPayload CreatePayload(string tag, string title,
            System.Collections.Generic.IEnumerable<string> targets,
            System.Collections.Generic.IEnumerable<string> parameters,
            System.Collections.Generic.IEnumerable<PreviewRowData> rows,
            System.Collections.Generic.IEnumerable<ParameterItem> parameterItems = null)
        {
            var cleanTargets = targets?
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList() ?? new System.Collections.Generic.List<string>();
            var cleanParameters = parameters?
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList() ?? new System.Collections.Generic.List<string>();
            var cleanRows = rows?
                .Where(r => r != null)
                .Take(PreviewValueHelpers.MaxRows)
                .ToList() ?? new System.Collections.Generic.List<PreviewRowData>();

            var kindMap = new System.Collections.Generic.Dictionary<string, ParamKind>(StringComparer.OrdinalIgnoreCase);
            if (parameterItems != null)
            {
                foreach (var p in parameterItems)
                {
                    if (p != null && !string.IsNullOrWhiteSpace(p.Name) && !kindMap.ContainsKey(p.Name))
                        kindMap[p.Name] = p.Kind;
                }
            }

            return new PreviewPayload(tag, title, cleanTargets, cleanParameters, cleanRows, kindMap);
        }

        private string ResolvePreviewSourceTag()
        {
            var active = GetActiveTabTag();
            if (!string.Equals(active, "Preview", StringComparison.OrdinalIgnoreCase))
                return active;
            return _previewPreferredTag ?? string.Empty;
        }

        private string GetActiveTabTag()
        {
            if (TabModel.IsChecked == true) return "Model";
            if (TabAnnotation.IsChecked == true) return "Annotation";
            if (TabElements.IsChecked == true) return "Elements";
            if (TabSchedules.IsChecked == true) return "Schedules";
            if (TabSpatial.IsChecked == true) return "Spatial";
            if (TabPreview.IsChecked == true) return "Preview";
            return string.Empty;
        }

        private void ComboBox_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var combo = FindParent<ComboBox>(e.OriginalSource as DependencyObject);
            if (combo == null || !combo.IsEnabled)
                return;

            if (!combo.IsDropDownOpen)
            {
                combo.Focus();
                combo.IsDropDownOpen = true;
                e.Handled = true;
            }
        }

        private void ComboBoxEditableText_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            var textBox = FindParent<TextBox>(e.OriginalSource as DependencyObject);
            if (textBox == null)
                return;

            var combo = FindParent<ComboBox>(textBox);
            if (combo == null)
                return;

            textBox.Cursor = Cursors.Hand;
            textBox.IsReadOnlyCaretVisible = false;
            textBox.Focusable = false;
        }

        private void ComboBox_MouseEnter(object sender, MouseEventArgs e)
        {
            var combo = FindParent<ComboBox>(e.OriginalSource as DependencyObject);
            if (combo == null || !combo.IsEnabled)
                return;

            if (!_comboDefaultVisuals.ContainsKey(combo))
                _comboDefaultVisuals[combo] = (combo.Background, combo.BorderBrush, combo.Effect);

            var hoverBg = new SolidColorBrush(Color.FromRgb(0xFF, 0xF3, 0xCC));
            var hoverBorder = new SolidColorBrush(Color.FromRgb(0xE8, 0xC4, 0x5A));
            combo.Background = hoverBg;
            combo.BorderBrush = hoverBorder;
            combo.Effect = new DropShadowEffect
            {
                Color = Color.FromRgb(0xB3, 0x8C, 0x2F),
                BlurRadius = 6,
                ShadowDepth = 1,
                Opacity = 0.35
            };

            combo.ApplyTemplate();
            if (combo.Template?.FindName("PART_EditableTextBox", combo) is TextBox editableText)
            {
                if (!_comboEditableDefaults.ContainsKey(editableText))
                    _comboEditableDefaults[editableText] = (editableText.Background, editableText.BorderBrush);

                editableText.Background = hoverBg;
                editableText.BorderBrush = hoverBorder;
            }

            var toggle = FindChild<ToggleButton>(combo);
            if (toggle != null)
            {
                if (!_comboToggleDefaults.ContainsKey(toggle))
                    _comboToggleDefaults[toggle] = (toggle.Background, toggle.BorderBrush);

                toggle.Background = hoverBg;
                toggle.BorderBrush = hoverBorder;
            }
        }

        private void ComboBox_MouseLeave(object sender, MouseEventArgs e)
        {
            var combo = FindParent<ComboBox>(e.OriginalSource as DependencyObject);
            if (combo == null)
                return;

            if (_comboDefaultVisuals.TryGetValue(combo, out var original))
            {
                combo.Background = original.Bg;
                combo.BorderBrush = original.Border;
                combo.Effect = original.Effect;
            }
            else
            {
                combo.Effect = null;
            }

            combo.ApplyTemplate();
            if (combo.Template?.FindName("PART_EditableTextBox", combo) is TextBox editableText)
            {
                if (_comboEditableDefaults.TryGetValue(editableText, out var tbOrig))
                {
                    editableText.Background = tbOrig.Bg;
                    editableText.BorderBrush = tbOrig.Border;
                }
                else
                {
                    editableText.ClearValue(BackgroundProperty);
                    editableText.ClearValue(BorderBrushProperty);
                }
            }

            var toggle = FindChild<ToggleButton>(combo);
            if (toggle != null)
            {
                if (_comboToggleDefaults.TryGetValue(toggle, out var tgOrig))
                {
                    toggle.Background = tgOrig.Bg;
                    toggle.BorderBrush = tgOrig.Border;
                }
                else
                {
                    toggle.ClearValue(BackgroundProperty);
                    toggle.ClearValue(BorderBrushProperty);
                }
            }
        }

        private static void ApplyComboBoxCursorMode(DependencyObject root)
        {
            if (root == null)
                return;

            if (root is ComboBox combo)
            {
                combo.Cursor = Cursors.Hand;
                combo.ApplyTemplate();
                if (combo.Template?.FindName("PART_EditableTextBox", combo) is TextBox tb)
                {
                    tb.Cursor = Cursors.Hand;
                    tb.IsReadOnlyCaretVisible = false;
                    tb.Focusable = false;
                }
            }

            var children = VisualTreeHelper.GetChildrenCount(root);
            for (var i = 0; i < children; i++)
                ApplyComboBoxCursorMode(VisualTreeHelper.GetChild(root, i));
        }

        private void ApplyButtonHoverMode(DependencyObject root)
        {
            if (root == null)
                return;

            if (root is Button button && _hoverButtons.Add(button))
            {
                button.Cursor = button.IsEnabled ? Cursors.Hand : Cursors.Arrow;
                button.RenderTransformOrigin = new Point(0.5, 0.5);
                button.MouseEnter += Button_MouseEnter;
                button.MouseLeave += Button_MouseLeave;
                button.IsEnabledChanged += Button_IsEnabledChanged;
            }

            var children = VisualTreeHelper.GetChildrenCount(root);
            for (var i = 0; i < children; i++)
                ApplyButtonHoverMode(VisualTreeHelper.GetChild(root, i));
        }

        private void Button_IsEnabledChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if (sender is not Button button)
                return;

            button.Cursor = button.IsEnabled ? Cursors.Hand : Cursors.Arrow;
            if (!button.IsEnabled)
                RestoreButtonHover(button);
        }

        private void Button_MouseEnter(object sender, MouseEventArgs e)
        {
            if (sender is not Button button || !button.IsEnabled)
                return;

            if (!_buttonDefaultVisuals.ContainsKey(button))
                _buttonDefaultVisuals[button] = (button.Background, button.BorderBrush, button.Effect, button.RenderTransform);

            button.Effect = new DropShadowEffect
            {
                Color = Color.FromRgb(0x5D, 0x4A, 0x1A),
                BlurRadius = 10,
                ShadowDepth = 3,
                Direction = 265,
                Opacity = 0.28
            };
            button.Background = new SolidColorBrush(Color.FromRgb(0xFF, 0xF3, 0xCC));
            button.BorderBrush = new SolidColorBrush(Color.FromRgb(0xE8, 0xC4, 0x5A));
            button.RenderTransform = new TranslateTransform(0, -1);
        }

        private void Button_MouseLeave(object sender, MouseEventArgs e)
        {
            if (sender is Button button)
                RestoreButtonHover(button);
        }

        private void RestoreButtonHover(Button button)
        {
            if (button == null)
                return;

            if (_buttonDefaultVisuals.TryGetValue(button, out var original))
            {
                button.Background = original.Bg;
                button.BorderBrush = original.Border;
                button.Effect = original.Effect;
                button.RenderTransform = original.Transform;
            }
            else
            {
                button.ClearValue(BackgroundProperty);
                button.ClearValue(BorderBrushProperty);
                button.Effect = null;
                button.RenderTransform = null;
            }
        }

        private static T FindChild<T>(DependencyObject root) where T : DependencyObject
        {
            if (root == null)
                return null;

            var childCount = VisualTreeHelper.GetChildrenCount(root);
            for (var i = 0; i < childCount; i++)
            {
                var child = VisualTreeHelper.GetChild(root, i);
                if (child is T found)
                    return found;

                var nested = FindChild<T>(child);
                if (nested != null)
                    return nested;
            }

            return null;
        }

        private static IEnumerable<T> EnumerateChildren<T>(DependencyObject root) where T : DependencyObject
        {
            if (root == null)
                yield break;

            var childCount = VisualTreeHelper.GetChildrenCount(root);
            for (var i = 0; i < childCount; i++)
            {
                var child = VisualTreeHelper.GetChild(root, i);
                if (child is T match)
                    yield return match;

                foreach (var nested in EnumerateChildren<T>(child))
                    yield return nested;
            }
        }

        private static T FindParent<T>(DependencyObject current) where T : DependencyObject
        {
            while (current != null)
            {
                if (current is T match)
                    return match;

                current = current switch
                {
                    FrameworkContentElement fce => fce.Parent,
                    Visual => VisualTreeHelper.GetParent(current),
                    System.Windows.Media.Media3D.Visual3D => VisualTreeHelper.GetParent(current),
                    _ => null
                };
            }

            return null;
        }

        // ══════════════════════════════════════════════════════════════════
        // PROGRESS (public — called from the views via ViewBase)
        // ══════════════════════════════════════════════════════════════════

        public void SetProgress(int pct, string msg = null)
            => _progress?.SetProgress(pct, msg);

        public void SetProgressError(string msg)
            => _progress?.SetError(msg);

        public void SetProgressWarning(string msg)
            => _progress?.SetWarning(msg);

        public void StartIndeterminateProgress(string msg = "Processing...")
            => _progress?.StartIndeterminate(msg);

        public void StopIndeterminateProgress()
            => _progress?.StopIndeterminate();

        // ── Busy overlay (public) ─────────────────────────────────────────

        public void ShowBusy(string msg = "Processing...")
        {
            _overlay?.Show(msg);
        }

        public void HideBusy()
        {
            _overlay?.Hide();
        }

        public void UpdateBusyMessage(string msg)
            => _overlay?.UpdateMessage(msg);

        // ══════════════════════════════════════════════════════════════════
        // PROFILE
        // ══════════════════════════════════════════════════════════════════

        private void LoadProfileList()
        {
            CbProfile.Items.Clear();
            var placeholder = new ComboBoxItem { Content = "Please Select" };
            CbProfile.Items.Add(placeholder);
            CbProfile.SelectedItem = placeholder;
            foreach (var name in _profileMgr.GetAllNames())
                CbProfile.Items.Add(new ComboBoxItem { Content = name });
        }

        private void CbProfile_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CbProfile.SelectedItem is not ComboBoxItem item) return;
            var name = item.Content?.ToString();
            if (string.IsNullOrEmpty(name) || name == "Please Select") return;
            var profile = _profileMgr.Load(name);
            if (profile == null) return;

            _profileMgr.ApplyToState(profile);
            ApplyProfileToViews(profile);
            _progress?.SetProgress(100, $"Profile \"{name}\" loaded.");
        }

        private void ApplyProfileToViews(ProfileData profile)
        {
            if (profile == null) return;
            try { ViewModelCategories?.ApplyProfile(profile); } catch { }
            try { ViewAnnotation?.ApplyProfile(profile); } catch { }
            try { ViewElements?.ApplyProfile(profile); } catch { }
            try { ViewSchedules?.ApplyProfile(profile); } catch { }
            try { ViewSpatial?.ApplyProfile(profile); } catch { }
            RefreshPreviewTabState();
        }

        private void BtnSaveProfile_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not Button btn)
            {
                SaveProfile();
                return;
            }

            var menu = new ContextMenu
            {
                PlacementTarget = btn,
                Placement = PlacementMode.Bottom
            };
            var miSave = new MenuItem { Header = "Save" };
            miSave.Click += (_, _) => SaveProfile();
            var miSaveAs = new MenuItem { Header = "Save As..." };
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
            if (string.IsNullOrWhiteSpace(name))
                return;

            try
            {
                var profile = BuildProfileSnapshotFromUi(name);
                _profileMgr.Save(profile);
                _profileMgr.ApplyToState(profile);
                LoadProfileList();
                SelectProfile(name);

                // Also save to Revit document extensible storage
                try
                {
                    var doc = Services.RevitDocumentCache.Current;
                    if (doc != null)
                        Services.RevitStorageService.SaveToDocument(doc, profile);
                }
                catch { }

                _progress?.SetProgress(100, $"Profile \"{name}\" saved.");
            }
            catch (Exception ex)
            {
                _progress?.SetError($"Error: {ex.Message}");
            }
        }

        private void SaveProfileAs()
        {
            var defaultName = GetCurrentProfileName();
            if (string.IsNullOrWhiteSpace(defaultName))
                defaultName = $"SheetLink_Profile_{DateTime.Now:yyyyMMdd_HHmm}";

            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Title = "Save Profile As",
                Filter = "Text Files|*.txt|All Files|*.*",
                FileName = defaultName + ".txt",
                AddExtension = true,
                DefaultExt = ".txt",
                OverwritePrompt = true
            };
            if (dlg.ShowDialog() != true)
                return;

            try
            {
                var fileName = Path.GetFileNameWithoutExtension(dlg.FileName);
                var profile = BuildProfileSnapshotFromUi(fileName);
                _profileMgr.ExportToFile(profile, dlg.FileName);

                // Also save to internal profiles and auto-select
                _profileMgr.Save(profile);
                _profileMgr.ApplyToState(profile);
                LoadProfileList();
                SelectProfile(fileName);

                // Save to Revit extensible storage
                try
                {
                    var doc = Services.RevitDocumentCache.Current;
                    if (doc != null)
                        Services.RevitStorageService.SaveToDocument(doc, profile);
                }
                catch { }

                _progress?.SetProgress(100, $"Profile saved: {fileName}");
            }
            catch (Exception ex)
            {
                _progress?.SetError($"Error: {ex.Message}");
            }
        }

        private string GetCurrentProfileName()
        {
            if (CbProfile.SelectedItem is ComboBoxItem item)
            {
                var fromCombo = item.Content?.ToString();
                if (!string.IsNullOrWhiteSpace(fromCombo) &&
                    !string.Equals(fromCombo, "Please Select", StringComparison.OrdinalIgnoreCase))
                {
                    return fromCombo;
                }
            }

            if (!string.IsNullOrWhiteSpace(_state.ActiveProfile?.Name) &&
                !string.Equals(_state.ActiveProfile.Name, "Please Select", StringComparison.OrdinalIgnoreCase))
            {
                return _state.ActiveProfile.Name;
            }

            return string.Empty;
        }

        private void SelectProfile(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return;

            foreach (ComboBoxItem item in CbProfile.Items)
            {
                if (string.Equals(item.Content?.ToString(), name, StringComparison.OrdinalIgnoreCase))
                {
                    CbProfile.SelectedItem = item;
                    return;
                }
            }
        }

        private ProfileData BuildProfileSnapshotFromUi(string name)
        {
            var profile = new ProfileData
            {
                Name = string.IsNullOrWhiteSpace(name) ? "New Profile" : name,
                LastExcelPath = _state.ActiveProfile?.LastExcelPath,
                LastGoogleDriveFolderId = _state.ActiveProfile?.LastGoogleDriveFolderId
            };

            profile.ModelCategories = ViewModelCategories?.GetPreviewTargets() ?? new List<string>();
            profile.ModelParameters = ViewModelCategories?.GetPreviewParameters() ?? new List<string>();
            profile.ModelScope = ViewModelCategories?.RbActiveView?.IsChecked == true
                ? "Active"
                : ViewModelCategories?.RbCurrentSelection?.IsChecked == true ? "Current" : "Whole";
            profile.ModelIncludeLinkedFiles = ViewModelCategories?.ChkLinked?.IsChecked == true;
            profile.ModelExportByTypeId = ViewModelCategories?.ChkByTypeId?.IsChecked == true;

            profile.AnnotationCategories = ViewAnnotation?.GetPreviewTargets() ?? new List<string>();
            profile.AnnotationParameters = ViewAnnotation?.GetPreviewParameters() ?? new List<string>();
            profile.AnnotationScope = ResolveScopeByGroupName(ViewAnnotation, "ScopeAnn");
            profile.AnnotationIncludeLinked = ViewAnnotation?.ChkLinked?.IsChecked == true;
            profile.AnnotationExportByTypeId = ViewAnnotation?.ChkByTypeId?.IsChecked == true;

            profile.ElementsCategory = ViewElements?.GetProfileSelectedCategory();
            profile.ElementsSelected = ViewElements?.GetProfileCheckedElements() ?? new List<string>();
            profile.ElementsParameters = ViewElements?.GetPreviewParameters() ?? new List<string>();
            profile.ElementsScope = ResolveScopeByGroupName(ViewElements, "ScopeElm");
            profile.ElementsIncludeLinked = ViewElements?.ChkLinked?.IsChecked == true;
            profile.ElementsExportByTypeId = ViewElements?.ChkByTypeId?.IsChecked == true;

            profile.Schedules = ViewSchedules?.GetProfileCheckedSchedules() ?? new List<string>();
            profile.ScheduleParameters = ViewSchedules?.GetPreviewParameters() ?? new List<string>();
            profile.ScheduleExportByTypeId = ViewSchedules?.GetProfileExportByTypeId() == true;
            profile.ScheduleScope = "Whole";

            profile.SpatialType = ViewSpatial?.GetProfileSpatialType() ?? "Rooms";
            profile.SpatialSelected = ViewSpatial?.GetProfileSelectedElementIds() ?? new List<long>();
            profile.SpatialParameters = ViewSpatial?.GetPreviewParameters() ?? new List<string>();
            profile.SpatialIncludeLinked = ViewSpatial?.GetProfileIncludeLinked() == true;

            return profile;
        }

        private static string ResolveScopeByGroupName(DependencyObject root, string groupName)
        {
            if (root == null || string.IsNullOrWhiteSpace(groupName))
                return "Whole";

            var selected = EnumerateChildren<RadioButton>(root)
                .FirstOrDefault(rb => rb?.IsChecked == true &&
                                      string.Equals(rb.GroupName, groupName, StringComparison.OrdinalIgnoreCase));
            var text = selected?.Content?.ToString() ?? string.Empty;
            if (text.IndexOf("Active", StringComparison.OrdinalIgnoreCase) >= 0)
                return "Active";
            if (text.IndexOf("Current", StringComparison.OrdinalIgnoreCase) >= 0)
                return "Current";
            return "Whole";
        }

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
                var profile = _profileMgr.LoadFromFile(dlg.FileName);
                if (profile == null)
                {
                    MessageBox.Show("The profile file is invalid.", "SheetLink", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (string.IsNullOrWhiteSpace(profile.Name))
                    profile.Name = Path.GetFileNameWithoutExtension(dlg.FileName);

                if (_profileMgr.Exists(profile.Name))
                {
                    var overwrite = MessageBox.Show(
                        $"Profile '{profile.Name}' already exists. Overwrite?",
                        "SheetLink",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Question);
                    if (overwrite != MessageBoxResult.Yes)
                        return;
                }

                _profileMgr.Save(profile);
                _profileMgr.ApplyToState(profile);
                LoadProfileList();
                SelectProfile(profile.Name);
                ApplyProfileToViews(profile);

                _progress?.SetProgress(100, $"Profile imported: {profile.Name}");
            }
            catch (Exception ex) { _progress?.SetError($"Error: {ex.Message}"); }
        }

        private void BtnDeleteProfile_Click(object sender, RoutedEventArgs e)
        {
            if (CbProfile.SelectedItem is not ComboBoxItem item) return;
            var name = item.Content?.ToString();
            if (string.IsNullOrEmpty(name) || name == "Please Select") return;
            var res = MessageBox.Show($"Delete profile \"{name}\"?",
                "SheetLink", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (res != MessageBoxResult.Yes) return;
            _profileMgr.Delete(name);
            LoadProfileList();
            _progress?.SetProgress(100, $"Profile \"{name}\" has been deleted.");
        }

        private void BtnAccount_Click(object sender, RoutedEventArgs e)
            => MessageBox.Show("THBIM account portal will be available soon.",
                "THBIM", MessageBoxButton.OK, MessageBoxImage.Information);

        private void BtnHelp_Click(object sender, RoutedEventArgs e)
            => MessageBox.Show("THBIM help center will be available soon.",
                "THBIM", MessageBoxButton.OK, MessageBoxImage.Information);

        private void BtnUpdate_Click(object sender, RoutedEventArgs e)
            => MessageBox.Show("THBIM update channel will be available soon.",
                "THBIM", MessageBoxButton.OK, MessageBoxImage.Information);

        // ══════════════════════════════════════════════════════════════════
        // KEYBOARD SHORTCUTS
        // ══════════════════════════════════════════════════════════════════

        private void RegisterShortcuts()
        {
            void Bind(Key key, Action action) =>
                InputBindings.Add(new KeyBinding(
                    new RelayCommand(_ => action()),
                    new KeyGesture(key, ModifierKeys.Control)));

            Bind(Key.D1, () => SwitchToTab("Model"));
            Bind(Key.D2, () => SwitchToTab("Annotation"));
            Bind(Key.D3, () => SwitchToTab("Elements"));
            Bind(Key.D4, () => SwitchToTab("Schedules"));
            Bind(Key.D5, () => SwitchToTab("Spatial"));
            Bind(Key.D6, () => SwitchToTab("Preview"));
            Bind(Key.S, SaveProfile);

            InputBindings.Add(new KeyBinding(
                new RelayCommand(_ =>
                {
                    if (_overlay?.Visibility != Visibility.Visible) Close();
                }),
                new KeyGesture(Key.Escape)));
        }

        // ══════════════════════════════════════════════════════════════════
        // INPUT DIALOG
        // ══════════════════════════════════════════════════════════════════

        private string ShowInputDialog(string prompt, string title,
                                        string defaultValue = "")
        {
            var win = new Window
            {
                Title = title,
                Width = 380,
                Height = 150,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = this,
                ResizeMode = ResizeMode.NoResize,
                WindowStyle = WindowStyle.ToolWindow,
                Background = new SolidColorBrush(Color.FromRgb(0xF4, 0xF4, 0xF4)),
                FontFamily = new FontFamily("Segoe UI"),
                ShowInTaskbar = false
            };

            var root = new Grid { Margin = new Thickness(16) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(6) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(10) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var lbl = new TextBlock
            {
                Text = prompt,
                FontSize = 12,
                Foreground = new SolidColorBrush(Color.FromRgb(0x1A, 0x1A, 0x1A))
            };
            Grid.SetRow(lbl, 0);

            var tb = new TextBox
            {
                Text = defaultValue,
                FontSize = 12,
                Height = 28,
                Padding = new Thickness(6, 4, 6, 4),
                BorderBrush = new SolidColorBrush(Color.FromRgb(0xD0, 0xD0, 0xD0)),
                BorderThickness = new Thickness(1),
                Background = Brushes.White
            };
            Grid.SetRow(tb, 2);

            var btnRow = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };
            Grid.SetRow(btnRow, 4);

            var btnOk = new Button
            {
                Content = "OK",
                Width = 76,
                Height = 28,
                Margin = new Thickness(0, 0, 8, 0),
                IsDefault = true,
                Background = new SolidColorBrush(Color.FromRgb(0x2A, 0x6E, 0x2A)),
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0),
                FontSize = 12,
                FontWeight = FontWeights.Medium,
                Cursor = Cursors.Hand
            };
            btnOk.Click += (_, _) => win.DialogResult = true;

            var btnCancel = new Button
            {
                Content = "Cancel",
                Width = 76,
                Height = 28,
                Background = new SolidColorBrush(Color.FromRgb(0xE0, 0xE0, 0xE0)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(0xD0, 0xD0, 0xD0)),
                BorderThickness = new Thickness(1),
                FontSize = 12,
                IsCancel = true,
                Cursor = Cursors.Hand
            };

            btnRow.Children.Add(btnOk);
            btnRow.Children.Add(btnCancel);
            root.Children.Add(lbl);
            root.Children.Add(tb);
            root.Children.Add(btnRow);
            win.Content = root;
            win.Loaded += (_, _) => { tb.Focus(); tb.SelectAll(); };

            return win.ShowDialog() == true ? tb.Text.Trim() : null;
        }

        private static void OpenUrl(string url)
        {
            try
            {
                System.Diagnostics.Process.Start(
                    new System.Diagnostics.ProcessStartInfo
                    { FileName = url, UseShellExecute = true });
            }
            catch { }
        }
    }
}








