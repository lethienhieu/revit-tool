using System.Windows;

namespace THBIM
{
    public partial class ProfileCreateDialog : Window
    {
        public string NewProfileName { get; private set; }
        public bool IsImportMode { get; private set; }

        public ProfileCreateDialog()
        {
            InitializeComponent();
            // Tự động focus vào ô nhập liệu để người dùng gõ được ngay
            TxtProfileName.Focus();
        }

        private void BtnCreate_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TxtProfileName.Text))
            {
                MessageBox.Show("Please enter a profile name.", "Notice");
                return;
            }

            NewProfileName = TxtProfileName.Text;
            IsImportMode = RbImport.IsChecked == true;

            this.DialogResult = true;
            this.Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }
    }
}