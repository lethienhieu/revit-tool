using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace THBIM.Models
{
    /// <summary>
    /// State trung tâm — chia sẻ giữa tất cả các tab.
    /// Singleton: SheetLinkState.Instance
    /// 
    /// Chứa:
    ///   - Danh sách categories / parameters / schedules / rooms đã load từ Revit
    ///   - Profile đang active
    ///   - Trạng thái progress bar
    /// </summary>
    public class SheetLinkState : INotifyPropertyChanged
    {
        // ── Singleton ────────────────────────────────────────────────────
        private static SheetLinkState _instance;
        public static SheetLinkState Instance => _instance ??= new SheetLinkState();
        private SheetLinkState() { }

        // ── Profile ───────────────────────────────────────────────────────
        private ProfileData _activeProfile = new();
        public ProfileData ActiveProfile
        {
            get => _activeProfile;
            set { _activeProfile = value; OnPropertyChanged(); }
        }

        // ── Model Categories ──────────────────────────────────────────────
        public ObservableCollection<CategoryItem> ModelCategories { get; } = new();
        public ObservableCollection<ParameterItem> ModelAvailParams { get; } = new();
        public ObservableCollection<ParameterItem> ModelSelectedParams { get; } = new();

        // ── Annotation Categories ─────────────────────────────────────────
        public ObservableCollection<CategoryItem> AnnCategories { get; } = new();
        public ObservableCollection<ParameterItem> AnnAvailParams { get; } = new();
        public ObservableCollection<ParameterItem> AnnSelectedParams { get; } = new();

        // ── Elements ──────────────────────────────────────────────────────
        public ObservableCollection<CategoryItem> ElemCategories { get; } = new();
        public ObservableCollection<CategoryItem> ElemElements { get; } = new();
        public ObservableCollection<ParameterItem> ElemAvailParams { get; } = new();
        public ObservableCollection<ParameterItem> ElemSelectedParams { get; } = new();

        // ── Schedules ─────────────────────────────────────────────────────
        public ObservableCollection<ScheduleItem> Schedules { get; } = new();
        public ObservableCollection<ParameterItem> SchedParams { get; } = new();

        // ── Spatial ───────────────────────────────────────────────────────
        public ObservableCollection<SpatialItem> SpatialItems { get; } = new();
        public ObservableCollection<ParameterItem> SpatialAvailParams { get; } = new();
        public ObservableCollection<ParameterItem> SpatialSelParams { get; } = new();

        // ── Progress ──────────────────────────────────────────────────────
        private int _progressPct;
        private string _progressMsg = "Completed   0%";
        private bool _isBusy;

        public int ProgressPct
        {
            get => _progressPct;
            set { _progressPct = value; OnPropertyChanged(); }
        }

        public string ProgressMsg
        {
            get => _progressMsg;
            set { _progressMsg = value; OnPropertyChanged(); }
        }

        public bool IsBusy
        {
            get => _isBusy;
            set { _isBusy = value; OnPropertyChanged(); }
        }

        // ── Helpers ───────────────────────────────────────────────────────

        /// <summary>Xóa sạch toàn bộ state về trạng thái ban đầu.</summary>
        public void Reset()
        {
            ModelCategories.Clear();
            ModelAvailParams.Clear();
            ModelSelectedParams.Clear();

            AnnCategories.Clear();
            AnnAvailParams.Clear();
            AnnSelectedParams.Clear();

            ElemCategories.Clear();
            ElemElements.Clear();
            ElemAvailParams.Clear();
            ElemSelectedParams.Clear();

            Schedules.Clear();
            SchedParams.Clear();

            SpatialItems.Clear();
            SpatialAvailParams.Clear();
            SpatialSelParams.Clear();

            ProgressPct = 0;
            ProgressMsg = "Completed   0%";
            IsBusy = false;
        }

        /// <summary>Cập nhật progress bar từ bất kỳ thread nào.</summary>
        public void SetProgress(int pct, string msg = null)
        {
            System.Windows.Application.Current?.Dispatcher.Invoke(() =>
            {
                ProgressPct = pct;
                ProgressMsg = msg ?? $"Completed   {pct}%";
            });
        }

        // ── INotifyPropertyChanged ────────────────────────────────────────
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string prop = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
    }
}