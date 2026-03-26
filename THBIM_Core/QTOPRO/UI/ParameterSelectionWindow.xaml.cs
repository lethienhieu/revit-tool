using System.Windows;

namespace THBIM
{
    public partial class ParameterSelectionWindow : Window
    {
        public ParameterSelectionWindow(ParameterSelectionViewModel viewModel)
        {
            InitializeComponent();
            this.DataContext = viewModel;
        }

        private void BtnOK_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
            this.Close();
        }
    }
}