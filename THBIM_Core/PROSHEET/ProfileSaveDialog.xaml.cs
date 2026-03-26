using System.Windows;

namespace THBIM
{
    public enum ProfileSaveResult { Cancel, Save, SaveAs }

    public partial class ProfileSaveDialog : Window
    {
        public ProfileSaveResult Result { get; private set; } = ProfileSaveResult.Cancel;

        public ProfileSaveDialog()
        {
            InitializeComponent();
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            Result = ProfileSaveResult.Save;
            this.DialogResult = true;
        }

        private void BtnSaveAs_Click(object sender, RoutedEventArgs e)
        {
            Result = ProfileSaveResult.SaveAs;
            this.DialogResult = true;
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
        }
    }
}