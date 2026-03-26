using System.Windows;

namespace colorslapsher.UI
{
    public partial class FilterPrefixWindow : Window
    {
        // Biến này để truyền dữ liệu về Form chính
        public string ResultPrefix { get; private set; } = "";

        public FilterPrefixWindow()
        {
            InitializeComponent();
            txtPrefix.Focus();
            txtPrefix.SelectAll();
        }

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            // Lấy giá trị nhập vào
            ResultPrefix = txtPrefix.Text.Trim();

            // Đóng window và báo thành công (True)
            this.DialogResult = true;
            this.Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }
    }
}