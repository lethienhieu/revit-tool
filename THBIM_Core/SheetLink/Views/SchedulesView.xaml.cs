using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using THBIM.Helpers;
using THBIM.Models;
using THBIM.Services;
using Microsoft.Win32;
using THBIM.Views;

namespace THBIM
{
    public partial class SchedulesView : UserControl
    {
        private readonly MockDataService _mock = new();
        private List<ScheduleItem> _all = new();
        private Border _selRow;
        private ScheduleItem _selectedSchedule;
        private static readonly SolidColorBrush BgSel = new(Color.FromRgb(0xDD, 0xEE, 0xFF));
        private static readonly SolidColorBrush BgNorm = Brushes.Transparent;
        private static readonly SolidColorBrush BgHover = new(Color.FromRgb(0xF0, 0xF6, 0xFF));

        public SchedulesView() { InitializeComponent(); Loaded += (_, _) => Load(); }

        // ── Load ──────────────────────────────────────────────────────────

        private void Load()
        {
            _all = ServiceLocator.IsRevitMode
                ? ServiceLocator.RevitData.GetSchedules()
                : _mock.GetSchedules();
            RenderList(_all);
            UpdateStatus();
        }

        private void RenderList(IEnumerable<ScheduleItem> items)
        {
            if (SpSchedules == null) return;

            SpSchedules.Children.Clear();
            if (items == null) return;

            foreach (var s in items.Where(x => x != null))
                SpSchedules.Children.Add(BuildRow(s));

            if (_selectedSchedule == null)
                return;

            _selRow = SafeFirstBorder(SpSchedules, _selectedSchedule);
            if (_selRow != null)
            {
                _selRow.Background = BgSel;
            }
            else
            {
                _selectedSchedule.IsSelected = false;
                _selectedSchedule = null;
                SpParams.Children.Clear();
            }
        }

        private Border BuildRow(ScheduleItem sched)
        {
            var row = new Border
            {
                Tag = sched,
                Height = 22,
                Padding = new Thickness(8, 0, 8, 0),
                Background = BgNorm,
                Cursor = Cursors.Hand
            };

            var sp = new StackPanel { Orientation = Orientation.Horizontal };
            var cb = new CheckBox
            {
                IsChecked = sched.IsChecked,
                Tag = sched,
                Margin = new Thickness(0, 0, 6, 0),
                VerticalContentAlignment = VerticalAlignment.Center
            };
            cb.Checked += ScheduleCheckBox_Changed;
            cb.Unchecked += ScheduleCheckBox_Changed;

            var tb = new TextBlock
            {
                Text = sched.Name,
                FontSize = 12.5,
                Foreground = Brushes.Black, // Use style constant
                VerticalAlignment = VerticalAlignment.Center,
                TextTrimming = TextTrimming.CharacterEllipsis
            };

            sp.Children.Add(cb);
            sp.Children.Add(tb);
            row.Child = sp;

            row.MouseEnter += (o, _) => { if ((Border)o != _selRow) ((Border)o).Background = BgHover; };
            row.MouseLeave += (o, _) => { if ((Border)o != _selRow) ((Border)o).Background = BgNorm; };
            row.MouseLeftButtonDown += Schedule_Click;

            return row;
        }

        // ── Click single-select ───────────────────────────────────────────

        private void Schedule_Click(object sender, MouseButtonEventArgs e)
        {
            if (sender is not Border row || row.Tag is not ScheduleItem sched) return;
            SelectSchedule(sched, row);
            UpdateStatus();
        }

        private void SelectSchedule(ScheduleItem sched, Border row = null)
        {
            if (_selRow != null)
            {
                _selRow.Background = BgNorm;
                if (_selRow.Tag is ScheduleItem prev) prev.IsSelected = false;
            }
            _selRow = row ?? SafeFirstBorder(SpSchedules, sched);
            if (_selRow != null)
                _selRow.Background = BgSel;

            _selectedSchedule = sched;
            sched.IsSelected = true;
            LoadScheduleParameters(sched);
        }

        private void LoadScheduleParameters(ScheduleItem sched)
        {
            var parms = ServiceLocator.IsRevitMode
                ? ServiceLocator.RevitData.GetScheduleParameters(sched.ElementId)
                : _mock.GetParameters(sched.Name);

            SpParams.Children.Clear();
            foreach (var p in parms)
                SpParams.Children.Add(ParamRowHelper.CreateRow(p));

            ApplyParamFilters();
        }

        private static void EnsureAtLeastOneChecked(CheckBox a, CheckBox b, CheckBox c)
        {
            if (a.IsChecked == true || b.IsChecked == true || c.IsChecked == true) return;
            a.IsChecked = true; b.IsChecked = true; c.IsChecked = true;
        }

        private static string UpdateFilterText(CheckBox inst, CheckBox typ, CheckBox ro, ComboBox cbFilter)
        {
            var parts = new List<string>();
            if (inst?.IsChecked == true) parts.Add("Instance");
            if (typ?.IsChecked == true) parts.Add("Type");
            if (ro?.IsChecked == true) parts.Add("Read-only");
            var text = string.Join(", ", parts);
            cbFilter.Text = text;
            return text;
        }

        private static bool PassKind(ParamKind kind, CheckBox inst, CheckBox typ, CheckBox ro)
            => kind switch
            {
                ParamKind.Instance => inst?.IsChecked == true,
                ParamKind.Type => typ?.IsChecked == true,
                ParamKind.ReadOnly => ro?.IsChecked == true,
                _ => false
            };

        private static IEnumerable<T> SafeChildren<T>(Panel panel) where T : UIElement
            => panel?.Children?.OfType<T>() ?? Enumerable.Empty<T>();

        private static Border SafeFirstBorder(Panel panel, object tag)
            => SafeChildren<Border>(panel).FirstOrDefault(b => ReferenceEquals(b?.Tag, tag));

        // ── Checkbox ─────────────────────────────────────────────────────

        private void ScheduleCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            if (sender is CheckBox cb && cb.Tag is ScheduleItem s)
            {
                s.IsChecked = cb.IsChecked == true;
                if (cb.IsChecked == true)
                {
                    var row = SafeFirstBorder(SpSchedules, s);
                    SelectSchedule(s, row);
                }
                else if (ReferenceEquals(_selectedSchedule, s))
                {
                    _selectedSchedule = null;
                    _selRow = null;
                    s.IsSelected = false;
                    SpParams.Children.Clear();
                }
            }

            int cnt = _all.Count(sc => sc.IsChecked);
            ChkAllSchedules.IsChecked = cnt == _all.Count ? true : cnt == 0 ? false : (bool?)null;
            UpdateStatus();
        }

        private void ScheduleCheckBox_Click(object sender, RoutedEventArgs e)
            => ScheduleCheckBox_Changed(sender, e);

        private void ChkAllSchedules_Click(object sender, RoutedEventArgs e)
        {
            bool check = ChkAllSchedules.IsChecked == true;
            foreach (var row in SafeChildren<Border>(SpSchedules).Where(r => r.Visibility == Visibility.Visible))
            {
                if (row.Child is StackPanel sp && sp.Children.OfType<CheckBox>().FirstOrDefault() is CheckBox cb)
                    cb.IsChecked = check;
            }
        }

        // ── Search / Filter ───────────────────────────────────────────────

        private void TbSrchSched_TextChanged(object sender, TextChangedEventArgs e)
        {
            ApplyScheduleFilters();
        }

        private void TbSrchParam_TextChanged(object sender, TextChangedEventArgs e)
        {
            ApplyParamFilters();
        }

        private void CbSchedType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!IsLoaded || SpSchedules == null) return;
            if (CbSchedType.SelectedItem is not ComboBoxItem item) return;
            var f = item.Content?.ToString();
            RenderList(f == "all" || string.IsNullOrEmpty(f)
                ? _all
                : _all.Where(sc => sc.KindLabel == f).ToList());
            UpdateStatus();
        }

        private void ChkHideUnchecked_Click(object sender, RoutedEventArgs e)
        {
            ApplyScheduleFilters();
        }

        private void ApplyScheduleFilters()
        {
            if (SpSchedules == null) return;

            var q = TbSrchSched.Text?.ToLower() ?? string.Empty;
            bool hide = ChkHideUnchecked?.IsChecked == true;

            foreach (Border row in SafeChildren<Border>(SpSchedules))
            {
                if (row.Tag is not ScheduleItem s) continue;
                var passSearch = string.IsNullOrEmpty(q) || s.Name.ToLower().Contains(q);
                var passHide = !hide || s.IsChecked;
                row.Visibility = passSearch && passHide ? Visibility.Visible : Visibility.Collapsed;
            }
            UpdateStatus();
        }

        private void ParamFilter_CheckedChanged(object sender, RoutedEventArgs e)
        {
            EnsureAtLeastOneChecked(ChkParamInst, ChkParamType, ChkParamRo);
            UpdateParamFilterText();
            ApplyParamFilters();
        }

        private void UpdateParamFilterText()
        {
            UpdateFilterText(ChkParamInst, ChkParamType, ChkParamRo, CbParamFilter);
        }

        private void ApplyParamFilters()
        {
            if (SpParams == null) return;

            var query = TbSrchParam?.Text?.Trim() ?? string.Empty;
            foreach (var row in SafeChildren<Border>(SpParams))
            {
                if (row.Tag is not ParameterItem p) continue;

                var passKind = PassKind(p.Kind, ChkParamInst, ChkParamType, ChkParamRo);
                var passText = string.IsNullOrWhiteSpace(query) || p.Name.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0;

                row.Visibility = passKind && passText ? Visibility.Visible : Visibility.Collapsed;
            }
            UpdateStatus();
        }

        private void BtnReset_Click(object sender, RoutedEventArgs e)
        {
            if (_selRow != null)
                _selRow.Background = BgNorm;
            _selRow = null;
            _selectedSchedule = null;
            foreach (Border row in SafeChildren<Border>(SpSchedules))
            {
                var cb = (row.Child as StackPanel)?.Children.OfType<CheckBox>().FirstOrDefault();
                cb?.Let(c => { c.IsChecked = false; if (c.Tag is ScheduleItem s) s.IsChecked = false; });
            }
            ChkAllSchedules.IsChecked = false;
            SpParams.Children.Clear();
            TbSrchSched.Clear();
            TbSrchParam.Clear();
            CbSchedType.SelectedIndex = 0;
            EnsureAtLeastOneChecked(ChkParamInst, ChkParamType, ChkParamRo);
            UpdateParamFilterText();
            UpdateStatus();
        }

        private void BtnPreview_Click(object sender, RoutedEventArgs e) => NavigateToPreview();

        private void BtnSectionBox_Click(object sender, RoutedEventArgs e)
            => MessageBox.Show("Section Box — Phase 3.", "SheetLink", MessageBoxButton.OK, MessageBoxImage.Information);

        private void BtnExportProjectStandards_Click(object sender, RoutedEventArgs e)
        {
            var opts = ShowProjectStandardsDialog();
            if (opts == null) return;
            if (opts.Google)
            {
                MessageBox.Show("Google export will be implemented later.", "THBIM", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            var selected = string.Join(", ", opts.Categories);
            if (string.IsNullOrWhiteSpace(selected)) selected = "None";
            MessageBox.Show(
                $"Project Standards export form is ready.\nSelected: {selected}\nOpen after export: {(opts.OpenAfter ? "Yes" : "No")}",
                "THBIM", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void BadgeInst_Click(object sender, MouseButtonEventArgs e)
        {
            ChkParamInst.IsChecked = true;
            ChkParamType.IsChecked = false;
            ChkParamRo.IsChecked = false;
            ParamFilter_CheckedChanged(sender, e);
        }

        private void BadgeType_Click(object sender, MouseButtonEventArgs e)
        {
            ChkParamInst.IsChecked = false;
            ChkParamType.IsChecked = true;
            ChkParamRo.IsChecked = false;
            ParamFilter_CheckedChanged(sender, e);
        }

        private void BadgeRo_Click(object sender, MouseButtonEventArgs e)
        {
            ChkParamInst.IsChecked = false;
            ChkParamType.IsChecked = false;
            ChkParamRo.IsChecked = true;
            ParamFilter_CheckedChanged(sender, e);
        }

        private void BtnImport_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog { Filter = "Excel|*.xlsx;*.xls|All|*.*" };
            if (dlg.ShowDialog() != true) return;
            if (!ServiceLocator.IsRevitMode) { ShowImportResult(new ImportResult(), "Import Schedules"); return; }

            var handler = Services.RevitEventHandler.Instance;
            if (handler == null) { MessageBox.Show("Revit event handler not initialized.", "THBIM", MessageBoxButton.OK, MessageBoxImage.Error); return; }

            var host = Window.GetWindow(this) as SheetLinkWindow;
            host?.ShowBusy("Importing from Excel...");

            ImportResult result = null;
            handler.Enqueue(
                _ =>
                {
                    result = ServiceLocator.Excel.ImportFromExcel(dlg.FileName, RevitDocumentCache.Current);
                },
                () => Dispatcher.Invoke(() =>
                {
                    host?.HideBusy();
                    result?.Let(r => ShowImportResult(r, "Import Schedules"));
                }));
        }

        private void CbImportAction_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!IsLoaded || sender is not ComboBox cb || cb.SelectedIndex <= 0)
                return;

            BtnImport_Click(sender, new RoutedEventArgs());
            cb.SelectedIndex = 0;
        }

        private async void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            var sel = _all.Where(sc => sc.IsChecked).ToList();
            if (!sel.Any()) { MessageBox.Show("Select at least one Schedule."); return; }
            var parms = GetPreviewParameters();

            var opts = ShowExportDialog();
            if (opts == null) return;
            if (opts.Google)
            {
                MessageBox.Show("Google export will be implemented later.", "THBIM", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var dlg = new SaveFileDialog { Filter = "Excel|*.xlsx", FileName = $"SheetLink_Schedules_{DateTime.Now:yyyyMMdd_HHmm}.xlsx" };
            if (dlg.ShowDialog() != true) return;

            await RunWithBusy($"Exporting {sel.Count} schedules...", async () =>
            {
                if (ServiceLocator.IsRevitMode)
                {
                    ServiceLocator.Excel.ExportSchedules(
                        dlg.FileName,
                        sel.Select(sc => sc.Name).ToList(),
                        parms,
                        ChkExportByTypeId?.IsChecked == true,
                        (_, msg) => Dispatcher.Invoke(() => UpdateBusyMessage(msg)));
                }
                await Task.CompletedTask;
            }, "Export completed.");

            if (opts.OpenAfter)
                OpenFile(dlg.FileName);
        }

        private void UpdateStatus()
        {
            int sel = _all.Count(sc => sc.IsChecked);
            int pFound = SafeChildren<Border>(SpParams).Count(b => b.Visibility == Visibility.Visible);
            TbStatus.Text = $"Selected schedules {sel} | parameters found {pFound}";
            BtnExport.IsEnabled = sel > 0;
            BtnPreview.IsEnabled = GetPreviewParameters().Any();
        }
    }
}

