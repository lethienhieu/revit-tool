using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace THBIM.Models
{
    /// <summary>
    /// Loại Schedule trong Revit.
    /// Dùng để filter dropdown trong tab Schedules.
    /// </summary>
    public enum ScheduleKind
    {
        Regular,        // Schedule thông thường
        KeySchedule,    // Key Schedule
        MaterialTakeoff,
        MultiCategory,
        PanelSchedule,
        SheetList,
        ViewList
    }

    /// <summary>
    /// Đại diện cho một Revit Schedule View.
    /// </summary>
    public class ScheduleItem : INotifyPropertyChanged
    {
        // ── Backing fields ──────────────────────────────────────────────
        private string _name;
        private ScheduleKind _kind;
        private bool _isChecked;
        private bool _isVisible = true;
        private bool _isSelected;   // click single-select để load params
        private string _elementId;

        // ── Properties ──────────────────────────────────────────────────

        /// <summary>Tên Schedule hiển thị trong UI (giống tên trong Revit Project Browser).</summary>
        public string Name
        {
            get => _name;
            set { _name = value; OnPropertyChanged(); }
        }

        /// <summary>Loại Schedule — dùng để filter dropdown "all / Key / Sheet List…"</summary>
        public ScheduleKind Kind
        {
            get => _kind;
            set { _kind = value; OnPropertyChanged(); OnPropertyChanged(nameof(KindLabel)); }
        }

        /// <summary>Label ngắn để hiển thị trong dropdown filter.</summary>
        public string KindLabel => Kind switch
        {
            ScheduleKind.KeySchedule => "Key Schedules",
            ScheduleKind.MaterialTakeoff => "Material Takeoff",
            ScheduleKind.MultiCategory => "Multi-Category",
            ScheduleKind.PanelSchedule => "Panel Schedule",
            ScheduleKind.SheetList => "Sheet List",
            ScheduleKind.ViewList => "View List",
            _ => "Regular"
        };

        /// <summary>Revit ElementId dạng string để tra cứu.</summary>
        public string ElementId
        {
            get => _elementId;
            set { _elementId = value; OnPropertyChanged(); }
        }

        /// <summary>Checkbox để chọn nhiều Schedule cùng lúc.</summary>
        public bool IsChecked
        {
            get => _isChecked;
            set { _isChecked = value; OnPropertyChanged(); }
        }

        /// <summary>
        /// Click đơn để load Parameters bên phải.
        /// Chỉ 1 schedule được selected tại một thời điểm.
        /// </summary>
        public bool IsSelected
        {
            get => _isSelected;
            set { _isSelected = value; OnPropertyChanged(); }
        }

        /// <summary>Ẩn/hiện khi search hoặc filter loại.</summary>
        public bool IsVisible
        {
            get => _isVisible;
            set { _isVisible = value; OnPropertyChanged(); }
        }

        // ── INotifyPropertyChanged ────────────────────────────────────────
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string prop = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));

        // ── Constructors ─────────────────────────────────────────────────
        public ScheduleItem() { }

        public ScheduleItem(string name, ScheduleKind kind = ScheduleKind.Regular,
                             string elementId = "")
        {
            _name = name;
            _kind = kind;
            _elementId = elementId;
        }

        public override string ToString() => Name ?? string.Empty;
    }
}