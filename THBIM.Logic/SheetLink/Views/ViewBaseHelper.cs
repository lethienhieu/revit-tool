using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using THBIM.Services;

namespace THBIM.Views
{
    internal static class ViewRuntimeHelpers
    {
        private static SheetLinkWindow GetHost(UserControl view)
            => Window.GetWindow(view) as SheetLinkWindow
               ?? Application.Current?.Windows.OfType<SheetLinkWindow>().FirstOrDefault();

        internal static async Task RunWithBusy(UserControl view, string busyMessage, Func<Task> action, string successMessage = null)
        {
            var host = GetHost(view);
            try
            {
                host?.ShowBusy(string.IsNullOrWhiteSpace(busyMessage) ? "Processing..." : busyMessage);
                if (action != null) await action();
                if (!string.IsNullOrWhiteSpace(successMessage))
                    host?.SetProgress(100, successMessage);
            }
            catch (Exception ex)
            {
                host?.SetProgressError(ex.Message);
                MessageBox.Show(ex.Message, "SheetLink", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                host?.HideBusy();
            }
        }

        internal static void UpdateBusyMessage(UserControl view, string message)
            => GetHost(view)?.UpdateBusyMessage(message ?? string.Empty);

        internal static void ShowImportResult(UserControl view, ImportResult result, string title)
        {
            if (result == null) return;
            var icon = result.Success ? MessageBoxImage.Information : MessageBoxImage.Warning;
            MessageBox.Show(result.Summary, title, MessageBoxButton.OK, icon);
        }

        internal static void NavigateToPreview(UserControl view)
        {
            var host = GetHost(view);
            if (host?.TabPreview == null) return;

            host.TryOpenPreview();
        }

        internal static void SetProgress(UserControl view, int percent, string message = null)
            => GetHost(view)?.SetProgress(percent, message);

        internal static SheetLinkWindow GetSheetLinkWindow(UserControl view)
            => GetHost(view);

        internal static void OpenFile(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return;
            var psi = new ProcessStartInfo(path) { UseShellExecute = true };
            Process.Start(psi);
        }

        // ── Shared Export Dialogs ──────────────────────────────────────────

        internal sealed class ExportOptions
        {
            public bool Google { get; set; }
            public bool OpenAfter { get; set; }
            public bool IncludeInstruction { get; set; }
        }

        internal sealed class ProjectStandardsOptions
        {
            public bool Google { get; set; }
            public bool OpenAfter { get; set; }
            public List<string> Categories { get; set; } = new();
        }

        internal static Window CreateDialogWindow(UserControl view, string title, int width, int height)
            => new()
            {
                Title = title,
                Width = width,
                Height = height,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = GetSheetLinkWindow(view),
                ResizeMode = ResizeMode.NoResize,
                WindowStyle = WindowStyle.ToolWindow,
                Background = new SolidColorBrush(Color.FromRgb(0xF4, 0xF4, 0xF4)),
                FontFamily = new System.Windows.Media.FontFamily("Segoe UI"),
                FontSize = 12
            };

        internal static void AddDialogLabel(Grid root, string text, int row, double size, bool bold)
        {
            var tb = new TextBlock
            {
                Text = text,
                FontSize = size,
                Margin = new Thickness(0, 0, 0, 10),
                Foreground = new SolidColorBrush(Color.FromRgb(0x2F, 0x2F, 0x2F)),
                FontWeight = bold ? FontWeights.Bold : FontWeights.Normal
            };
            Grid.SetRow(tb, row);
            root.Children.Add(tb);
        }

        internal static ExportOptions ShowExportDialog(UserControl view)
        {
            var win = CreateDialogWindow(view, "THBIM - Export", 420, 260);
            var root = new Grid { Margin = new Thickness(16) };
            for (var i = 0; i < 5; i++) root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

            AddDialogLabel(root, "Export Options", 0, 18, true);
            var chkOpen = new CheckBox { Content = "Open Excel file after export", IsChecked = true, Margin = new Thickness(0, 0, 0, 6) };
            Grid.SetRow(chkOpen, 1); root.Children.Add(chkOpen);
            var chkInstr = new CheckBox { Content = "Include Instructions sheet", IsChecked = true, Margin = new Thickness(0, 0, 0, 14) };
            Grid.SetRow(chkInstr, 2); root.Children.Add(chkInstr);

            var btnRow = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            var btnExcel = new Button { Content = "Export to Excel", Width = 130, Height = 32, Margin = new Thickness(0, 0, 8, 0) };
            var btnGoogle = new Button { Content = "Export to Google", Width = 130, Height = 32 };
            btnRow.Children.Add(btnExcel); btnRow.Children.Add(btnGoogle);
            Grid.SetRow(btnRow, 4); root.Children.Add(btnRow);

            ExportOptions result = null;
            btnExcel.Click += (_, _) =>
            {
                result = new ExportOptions { Google = false, OpenAfter = chkOpen.IsChecked == true, IncludeInstruction = chkInstr.IsChecked == true };
                win.DialogResult = true;
            };
            btnGoogle.Click += (_, _) =>
            {
                result = new ExportOptions { Google = true, OpenAfter = false, IncludeInstruction = chkInstr.IsChecked == true };
                win.DialogResult = true;
            };

            win.Content = root;
            return win.ShowDialog() == true ? result : null;
        }

        internal static ProjectStandardsOptions ShowProjectStandardsDialog(UserControl view)
        {
            var win = CreateDialogWindow(view, "THBIM - Export Project Standards", 520, 430);
            var root = new Grid { Margin = new Thickness(16) };
            for (var i = 0; i < 6; i++) root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

            AddDialogLabel(root, "THBIM", 0, 30, true);
            var grpExport = new GroupBox { Header = "Export Options", Margin = new Thickness(0, 0, 0, 10) };
            var chkOpen = new CheckBox { Content = "Open Excel File After Export", IsChecked = true, Margin = new Thickness(10, 6, 10, 8) };
            grpExport.Content = chkOpen;
            Grid.SetRow(grpExport, 1); root.Children.Add(grpExport);

            var grpChoose = new GroupBox { Header = "Choose Categories", Margin = new Thickness(0, 0, 0, 14) };
            var panel = new StackPanel { Margin = new Thickness(10, 6, 10, 8) };
            var chkProjectInfo = new CheckBox { Content = "Project Information", IsChecked = true };
            var chkObjectStyles = new CheckBox { Content = "Object Styles", IsChecked = true };
            var chkLineStyles = new CheckBox { Content = "Line Styles", IsChecked = true };
            var chkFamilies = new CheckBox { Content = "Families", IsChecked = true };
            panel.Children.Add(chkProjectInfo); panel.Children.Add(chkObjectStyles);
            panel.Children.Add(chkLineStyles); panel.Children.Add(chkFamilies);
            grpChoose.Content = panel;
            Grid.SetRow(grpChoose, 2); root.Children.Add(grpChoose);

            var btnRow = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Center };
            var btnExcel = new Button { Content = "Excel", Width = 130, Height = 42, Margin = new Thickness(0, 0, 10, 0) };
            var btnGoogle = new Button { Content = "Google", Width = 130, Height = 42 };
            btnRow.Children.Add(btnExcel); btnRow.Children.Add(btnGoogle);
            Grid.SetRow(btnRow, 3); root.Children.Add(btnRow);

            ProjectStandardsOptions result = null;
            void Capture(bool google)
            {
                var list = new List<string>();
                if (chkProjectInfo.IsChecked == true) list.Add("Project Information");
                if (chkObjectStyles.IsChecked == true) list.Add("Object Styles");
                if (chkLineStyles.IsChecked == true) list.Add("Line Styles");
                if (chkFamilies.IsChecked == true) list.Add("Families");
                result = new ProjectStandardsOptions { Google = google, OpenAfter = chkOpen.IsChecked == true, Categories = list };
                win.DialogResult = true;
            }
            btnExcel.Click += (_, _) => Capture(false);
            btnGoogle.Click += (_, _) => Capture(true);

            win.Content = root;
            return win.ShowDialog() == true ? result : null;
        }
    }
}

namespace THBIM
{
    public partial class AnnotationCategoriesView
    {
        private Task RunWithBusy(string busyMessage, Func<Task> action, string successMessage = null)
            => Views.ViewRuntimeHelpers.RunWithBusy(this, busyMessage, action, successMessage);

        private void UpdateBusyMessage(string message)
            => Views.ViewRuntimeHelpers.UpdateBusyMessage(this, message);

        private void ShowImportResult(ImportResult result, string title)
            => Views.ViewRuntimeHelpers.ShowImportResult(this, result, title);

        private void NavigateToPreview()
            => Views.ViewRuntimeHelpers.NavigateToPreview(this);

        private SheetLinkWindow GetSheetLinkWindow()
            => Views.ViewRuntimeHelpers.GetSheetLinkWindow(this);

        private void OpenFile(string path)
            => Views.ViewRuntimeHelpers.OpenFile(path);

        private Views.ViewRuntimeHelpers.ExportOptions ShowExportDialog()
            => Views.ViewRuntimeHelpers.ShowExportDialog(this);

        private Views.ViewRuntimeHelpers.ProjectStandardsOptions ShowProjectStandardsDialog()
            => Views.ViewRuntimeHelpers.ShowProjectStandardsDialog(this);
    }

    public partial class ElementsView
    {
        private Task RunWithBusy(string busyMessage, Func<Task> action, string successMessage = null)
            => Views.ViewRuntimeHelpers.RunWithBusy(this, busyMessage, action, successMessage);

        private void UpdateBusyMessage(string message)
            => Views.ViewRuntimeHelpers.UpdateBusyMessage(this, message);

        private void ShowImportResult(ImportResult result, string title)
            => Views.ViewRuntimeHelpers.ShowImportResult(this, result, title);

        private void NavigateToPreview()
            => Views.ViewRuntimeHelpers.NavigateToPreview(this);

        private SheetLinkWindow GetSheetLinkWindow()
            => Views.ViewRuntimeHelpers.GetSheetLinkWindow(this);

        private void OpenFile(string path)
            => Views.ViewRuntimeHelpers.OpenFile(path);

        private Views.ViewRuntimeHelpers.ExportOptions ShowExportDialog()
            => Views.ViewRuntimeHelpers.ShowExportDialog(this);

        private Views.ViewRuntimeHelpers.ProjectStandardsOptions ShowProjectStandardsDialog()
            => Views.ViewRuntimeHelpers.ShowProjectStandardsDialog(this);
    }

    public partial class ModelCategoriesView
    {
        private Task RunWithBusy(string busyMessage, Func<Task> action, string successMessage = null)
            => Views.ViewRuntimeHelpers.RunWithBusy(this, busyMessage, action, successMessage);

        private void UpdateBusyMessage(string message)
            => Views.ViewRuntimeHelpers.UpdateBusyMessage(this, message);

        private void ShowImportResult(ImportResult result, string title)
            => Views.ViewRuntimeHelpers.ShowImportResult(this, result, title);

        private void NavigateToPreview()
            => Views.ViewRuntimeHelpers.NavigateToPreview(this);

        private void SetProgress(int percent, string message = null)
            => Views.ViewRuntimeHelpers.SetProgress(this, percent, message);

        private SheetLinkWindow GetSheetLinkWindow()
            => Views.ViewRuntimeHelpers.GetSheetLinkWindow(this);

        private void OpenFile(string path)
            => Views.ViewRuntimeHelpers.OpenFile(path);
    }

    public partial class SchedulesView
    {
        private Task RunWithBusy(string busyMessage, Func<Task> action, string successMessage = null)
            => Views.ViewRuntimeHelpers.RunWithBusy(this, busyMessage, action, successMessage);

        private void UpdateBusyMessage(string message)
            => Views.ViewRuntimeHelpers.UpdateBusyMessage(this, message);

        private void ShowImportResult(ImportResult result, string title)
            => Views.ViewRuntimeHelpers.ShowImportResult(this, result, title);

        private void NavigateToPreview()
            => Views.ViewRuntimeHelpers.NavigateToPreview(this);

        private SheetLinkWindow GetSheetLinkWindow()
            => Views.ViewRuntimeHelpers.GetSheetLinkWindow(this);

        private void OpenFile(string path)
            => Views.ViewRuntimeHelpers.OpenFile(path);

        private Views.ViewRuntimeHelpers.ExportOptions ShowExportDialog()
            => Views.ViewRuntimeHelpers.ShowExportDialog(this);

        private Views.ViewRuntimeHelpers.ProjectStandardsOptions ShowProjectStandardsDialog()
            => Views.ViewRuntimeHelpers.ShowProjectStandardsDialog(this);
    }

    public partial class SpatialView
    {
        private Task RunWithBusy(string busyMessage, Func<Task> action, string successMessage = null)
            => Views.ViewRuntimeHelpers.RunWithBusy(this, busyMessage, action, successMessage);

        private void UpdateBusyMessage(string message)
            => Views.ViewRuntimeHelpers.UpdateBusyMessage(this, message);

        private void ShowImportResult(ImportResult result, string title)
            => Views.ViewRuntimeHelpers.ShowImportResult(this, result, title);

        private void NavigateToPreview()
            => Views.ViewRuntimeHelpers.NavigateToPreview(this);

        private SheetLinkWindow GetSheetLinkWindow()
            => Views.ViewRuntimeHelpers.GetSheetLinkWindow(this);

        private void OpenFile(string path)
            => Views.ViewRuntimeHelpers.OpenFile(path);

        private Views.ViewRuntimeHelpers.ExportOptions ShowExportDialog()
            => Views.ViewRuntimeHelpers.ShowExportDialog(this);
    }
}
