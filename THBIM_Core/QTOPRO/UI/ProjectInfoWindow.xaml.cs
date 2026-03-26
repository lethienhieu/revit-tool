using System.Windows;

namespace THBIM
{
    public partial class ProjectInfoWindow : Window
    {
        public ProjectInfo ResultInfo { get; private set; }

        public ProjectInfoWindow(ProjectInfo currentInfo = null)
        {
            InitializeComponent();
            if (currentInfo != null)
            {
                txtProjectName.Text = currentInfo.ProjectName;
                txtMeasuredBy.Text = currentInfo.MeasuredBy;
                txtRevision.Text = currentInfo.Revision;
            }
        }

        private void BtnOK_Click(object sender, RoutedEventArgs e)
        {
            ResultInfo = new ProjectInfo
            {
                ProjectName = txtProjectName.Text,
                MeasuredBy = txtMeasuredBy.Text,
                Revision = txtRevision.Text
            };
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