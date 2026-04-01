using Autodesk.Revit.UI;
using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Threading;

namespace GridBubble
{
    public enum BubbleMode { Hide, Show }

    public partial class GridBubbleWindow : Window
    {
        
        private readonly UIApplication _uiapp;
        private readonly GridBubbleExternalHandler _handler;
        private readonly ExternalEvent _externalEvent;
        private readonly IntPtr _revitHwnd;
        private bool _ready = false;
        private const double RIBBON_TITLE_BAND_OFFSET_PX = 153;
        private const double LEFT_MARGIN_PX = 400;

        public GridBubbleWindow(ExternalCommandData commandData)
        {
            _uiapp = commandData.Application;
            _handler = new GridBubbleExternalHandler(this);
            _externalEvent = ExternalEvent.Create(_handler);

            InitializeComponent();

            _revitHwnd = _uiapp.MainWindowHandle;
            new WindowInteropHelper(this) { Owner = _revitHwnd };

            Loaded += (_, __) => PositionUnderRibbonPanelName();
            _ready = true;
        }

        // ... (Giữ nguyên OnSourceInitialized, PositionUnderRibbonPanelName, Win32 API...) ...
        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            IntPtr hwnd = new WindowInteropHelper(this).Handle;
            if (hwnd != IntPtr.Zero)
            {
                int ex = GetWindowLong(hwnd, GWL_EXSTYLE);
                ex |= WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW;
                SetWindowLong(hwnd, GWL_EXSTYLE, ex);
            }
        }

        // ... (Giữ nguyên các hàm bổ trợ cũ) ...

        private void StartOrContinue(BubbleMode mode)
        {
            if (!_ready) return;
            // Cập nhật Mode cho handler (để dự phòng)
            _handler.Mode = mode;

            GiveFocusHardToRevit();

            if (!_handler.SessionRunning)
            {
                Dispatcher.BeginInvoke(
                    new Action(() => _externalEvent.Raise()),
                    DispatcherPriority.ApplicationIdle);
            }
            // ... (Phần disable hit test giữ nguyên) ...
            try { Mouse.Capture(null); } catch { }
            try { Keyboard.ClearFocus(); } catch { }

            IsHitTestVisible = false;
            var t = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(60) };
            t.Tick += (s, e) =>
            {
                ((DispatcherTimer)s!).Stop();
                IsHitTestVisible = true;
            };
            t.Start();
        }

        // ===== EVENT HANDLERS (Giữ nguyên) =====
        private void RbHide_PreviewMouseLeftButtonDown(object s, MouseButtonEventArgs e)
            => StartOrContinue(BubbleMode.Hide);

        private void RbShow_PreviewMouseLeftButtonDown(object s, MouseButtonEventArgs e)
            => StartOrContinue(BubbleMode.Show);

        private void RbHide_Checked(object s, RoutedEventArgs e) { }
        private void RbShow_Checked(object s, RoutedEventArgs e) { }
        private void RbHide_Click(object s, RoutedEventArgs e) { }
        private void RbShow_Click(object s, RoutedEventArgs e) { }


        // ===== FIX CHÍNH: HÀM LẤY TRẠNG THÁI TRỰC TIẾP TỪ UI =====
        // Hàm này bắt buộc chạy trên UI Thread để đọc IsChecked chuẩn xác nhất
        internal BubbleMode GetModeDirectly()
        {
            return this.Dispatcher.Invoke(() =>
            {
                if (RbHide.IsChecked == true) return BubbleMode.Hide;
                return BubbleMode.Show;
            });
        }

        internal void EndSessionAndClose()
        {
            if (Dispatcher.CheckAccess()) Close();
            else Dispatcher.Invoke(Close);
        }

        // ... (Giữ nguyên phần Win32 API ở dưới cùng) ...
        private const int GWL_EXSTYLE = -20;
        private const int WS_EX_TOOLWINDOW = 0x00000080;
        private const int WS_EX_NOACTIVATE = 0x08000000;

        [DllImport("user32.dll")] private static extern int GetWindowLong(IntPtr hWnd, int nIndex);
        [DllImport("user32.dll")] private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
        [DllImport("user32.dll", SetLastError = true)] private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern IntPtr SetFocus(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern IntPtr SetActiveWindow(IntPtr hWnd);
        [DllImport("user32.dll", CharSet = CharSet.Auto)] private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string? lpszClass, string? lpszWindow);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")] private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);
        [DllImport("user32.dll", CharSet = CharSet.Auto)] private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
        [StructLayout(LayoutKind.Sequential)] private struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
        // ... (Kết thúc Win32) ...
        private void PositionUnderRibbonPanelName()
        {
            if (_revitHwnd == IntPtr.Zero) return;
            if (!GetWindowRect(_revitHwnd, out RECT r)) return;
            double dpi = GetDpiScale(_revitHwnd);
            Left = (r.Left / dpi) + LEFT_MARGIN_PX;
            Top = (r.Top / dpi) + RIBBON_TITLE_BAND_OFFSET_PX;
        }
        private static double GetDpiScale(IntPtr owner)
        {
            HwndSource? src = HwndSource.FromHwnd(owner);
            return src?.CompositionTarget?.TransformToDevice.M11 ?? 1.0;
        }
        private void GiveFocusHardToRevit()
        {
            if (_revitHwnd == IntPtr.Zero) return;
            uint revitTid = GetWindowThreadProcessId(_revitHwnd, out _);
            uint thisTid = GetCurrentThreadId();
            AttachThreadInput(thisTid, revitTid, true);
            SetForegroundWindow(_revitHwnd);
            IntPtr mdi = FindWindowEx(_revitHwnd, IntPtr.Zero, "MDIClient", null);
            if (mdi != IntPtr.Zero)
            {
                const int WM_MDIGETACTIVE = 0x0229;
                IntPtr activeChild = SendMessage(mdi, WM_MDIGETACTIVE, IntPtr.Zero, IntPtr.Zero);
                if (activeChild != IntPtr.Zero) { SetActiveWindow(activeChild); SetFocus(activeChild); }
                else { SetActiveWindow(mdi); SetFocus(mdi); }
            }
            else { SetActiveWindow(_revitHwnd); SetFocus(_revitHwnd); }
            AttachThreadInput(thisTid, revitTid, false);
        }
    }
}