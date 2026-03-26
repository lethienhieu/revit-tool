using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace THBIM.Models
{
    /// <summary>
    /// Đại diện cho một Room hoặc Space trong Revit.
    /// Dùng trong tab Spatial — hiển thị dạng bảng Number | Name.
    /// </summary>
    public class SpatialItem : INotifyPropertyChanged
    {
        // ── Backing fields ──────────────────────────────────────────────
        private string _number;
        private string _name;
        private string _phase;
        private bool _isChecked;
        private bool _isVisible = true;
        private long _elementId;
        private bool _isRoom = true;   // false = Space

        // ── Properties ──────────────────────────────────────────────────

        /// <summary>Số phòng (Room Number). Bắt buộc khi tạo mới từ Excel.</summary>
        public string Number
        {
            get => _number;
            set { _number = value; OnPropertyChanged(); }
        }

        /// <summary>Tên phòng (Room Name). Bắt buộc khi tạo mới từ Excel.</summary>
        public string Name
        {
            get => _name;
            set { _name = value; OnPropertyChanged(); }
        }

        /// <summary>
        /// Phase mà Room/Space thuộc về.
        /// Bắt buộc khi tạo mới từ Excel.
        /// </summary>
        public string Phase
        {
            get => _phase;
            set { _phase = value; OnPropertyChanged(); }
        }

        /// <summary>
        /// Revit Element ID. 
        /// -1 = chưa tồn tại trong model (hàng mới trong Excel cần tạo).
        /// </summary>
        public long ElementId
        {
            get => _elementId;
            set { _elementId = value; OnPropertyChanged(); OnPropertyChanged(nameof(IsNew)); }
        }

        /// <summary>True nếu chưa có ElementId — cần tạo mới khi import.</summary>
        public bool IsNew => ElementId <= 0;

        /// <summary>True = Room, False = Space.</summary>
        public bool IsRoom
        {
            get => _isRoom;
            set { _isRoom = value; OnPropertyChanged(); OnPropertyChanged(nameof(TypeLabel)); }
        }

        /// <summary>"Room" hoặc "Space" — hiện ở dropdown Spatial toolbar.</summary>
        public string TypeLabel => IsRoom ? "Room" : "Space";

        /// <summary>Trạng thái checkbox trong danh sách Select Rooms/Spaces.</summary>
        public bool IsChecked
        {
            get => _isChecked;
            set { _isChecked = value; OnPropertyChanged(); }
        }

        /// <summary>Ẩn/hiện khi search.</summary>
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
        public SpatialItem() { }

        public SpatialItem(long elementId, string number, string name,
                           string phase, bool isRoom = true)
        {
            _elementId = elementId;
            _number = number;
            _name = name;
            _phase = phase;
            _isRoom = isRoom;
        }

        public override string ToString() => $"{Number} - {Name}";
    }
}