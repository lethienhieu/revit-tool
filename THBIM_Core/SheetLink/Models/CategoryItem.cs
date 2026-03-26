using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace THBIM.Models
{
    /// <summary>
    /// Đại diện cho một Revit Category (Model hoặc Annotation).
    /// Implement INotifyPropertyChanged để binding với WPF tự động cập nhật UI.
    /// </summary>
    public class CategoryItem : INotifyPropertyChanged
    {
        // ── Backing fields ──────────────────────────────────────────────
        private string _name;
        private string _discipline;
        private bool _isChecked;
        private bool _isVisible = true;

        // ── Properties ──────────────────────────────────────────────────

        /// <summary>Tên Category trong Revit (ví dụ: "Walls", "Doors").</summary>
        public string Name
        {
            get => _name;
            set { _name = value; OnPropertyChanged(); }
        }

        /// <summary>
        /// Discipline phân loại: "Architecture", "Structure", "MEP", hoặc "Other".
        /// Dùng để filter dropdown &lt;All Disciplines&gt;.
        /// </summary>
        public string Discipline
        {
            get => _discipline;
            set { _discipline = value; OnPropertyChanged(); }
        }

        /// <summary>Trạng thái checkbox trong danh sách Select Categories.</summary>
        public bool IsChecked
        {
            get => _isChecked;
            set
            {
                _isChecked = value;
                OnPropertyChanged();
                // Notify parent nếu cần cập nhật status bar
                CheckedChanged?.Invoke(this, value);
            }
        }

        /// <summary>
        /// Dùng để ẩn/hiện row khi người dùng search hoặc bật "Hide un-checked".
        /// </summary>
        public bool IsVisible
        {
            get => _isVisible;
            set { _isVisible = value; OnPropertyChanged(); }
        }

        // ── Event ────────────────────────────────────────────────────────
        /// <summary>Raised khi IsChecked thay đổi. Tham số bool = giá trị mới.</summary>
        public event System.Action<CategoryItem, bool> CheckedChanged;

        // ── INotifyPropertyChanged ────────────────────────────────────────
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        // ── Constructors ─────────────────────────────────────────────────
        public CategoryItem() { }

        public CategoryItem(string name, string discipline = "Other")
        {
            _name = name;
            _discipline = discipline;
        }

        public override string ToString() => Name ?? string.Empty;
    }
}