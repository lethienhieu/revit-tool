using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Media;

namespace THBIM.Models
{
    /// <summary>
    /// Loại parameter — dùng để tô màu background trong danh sách Available/Selected.
    /// </summary>
    public enum ParamKind
    {
        Instance,   // xanh lá  #EAF3DE
        Type,       // vàng     #FDF0CC
        ReadOnly    // đỏ nhạt  #FDEAEA
    }

    /// <summary>
    /// Đại diện cho một Revit Parameter.
    /// Chứa đủ thông tin để hiển thị trong UI và để export/import dữ liệu.
    /// </summary>
    public class ParameterItem : INotifyPropertyChanged
    {
        // ── Backing fields ──────────────────────────────────────────────
        private string _name;
        private string _storageType;
        private ParamKind _kind;
        private bool _isHighlighted;
        private bool _isVisible = true;

        // ── Properties ──────────────────────────────────────────────────

        /// <summary>Tên parameter hiển thị trong UI và dùng làm header Excel.</summary>
        public string Name
        {
            get => _name;
            set { _name = value; OnPropertyChanged(); }
        }

        /// <summary>
        /// Kiểu dữ liệu Revit: "String", "Double", "Integer", "ElementId".
        /// Hiển thị ở dòng thứ 2 trong header của PreviewEditView.
        /// </summary>
        public string StorageType
        {
            get => _storageType;
            set { _storageType = value; OnPropertyChanged(); }
        }

        /// <summary>Instance / Type / ReadOnly — quyết định màu background row.</summary>
        public ParamKind Kind
        {
            get => _kind;
            set
            {
                _kind = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(RowBackground));
                OnPropertyChanged(nameof(DotColor));
                OnPropertyChanged(nameof(KindLabel));
            }
        }

        /// <summary>Row đang được highlight (chọn để move ◀ ▶).</summary>
        public bool IsHighlighted
        {
            get => _isHighlighted;
            set { _isHighlighted = value; OnPropertyChanged(); OnPropertyChanged(nameof(RowBackground)); }
        }

        /// <summary>Ẩn/hiện khi search trong Available / Selected Parameters.</summary>
        public bool IsVisible
        {
            get => _isVisible;
            set { _isVisible = value; OnPropertyChanged(); }
        }

        // ── Computed display properties ──────────────────────────────────

        /// <summary>Màu nền row theo Kind. Khi highlight thêm border xanh.</summary>
        public Brush RowBackground
        {
            get
            {
                if (IsHighlighted)
                    return new SolidColorBrush(Color.FromRgb(0xCC, 0xE5, 0xFF));

                return Kind switch
                {
                    ParamKind.Instance => new SolidColorBrush(Color.FromRgb(0xEA, 0xF3, 0xDE)),
                    ParamKind.Type => new SolidColorBrush(Color.FromRgb(0xFD, 0xF0, 0xCC)),
                    ParamKind.ReadOnly => new SolidColorBrush(Color.FromRgb(0xFD, 0xEA, 0xEA)),
                    _ => Brushes.White
                };
            }
        }

        /// <summary>Màu chấm tròn nhỏ bên phải tên parameter.</summary>
        public Brush DotColor => Kind switch
        {
            ParamKind.Instance => new SolidColorBrush(Color.FromRgb(0x7A, 0xB8, 0x7A)),
            ParamKind.Type => new SolidColorBrush(Color.FromRgb(0xD4, 0xAA, 0x40)),
            ParamKind.ReadOnly => new SolidColorBrush(Color.FromRgb(0xD0, 0x70, 0x70)),
            _ => Brushes.Gray
        };

        /// <summary>Label ngắn gọn: "Instance" / "Type" / "Read-only".</summary>
        public string KindLabel => Kind switch
        {
            ParamKind.Instance => "Instance",
            ParamKind.Type => "Type",
            ParamKind.ReadOnly => "Read-only",
            _ => string.Empty
        };

        // ── INotifyPropertyChanged ────────────────────────────────────────
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        // ── Constructors ─────────────────────────────────────────────────
        public ParameterItem() { }

        public ParameterItem(string name, string storageType, ParamKind kind)
        {
            _name = name;
            _storageType = storageType;
            _kind = kind;
        }

        public override string ToString() => Name ?? string.Empty;
    }
}