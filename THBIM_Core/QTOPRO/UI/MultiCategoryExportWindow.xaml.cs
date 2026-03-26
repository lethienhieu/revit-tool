using Autodesk.Revit.DB;
using System.Windows;

namespace THBIM
{
    public partial class MultiCategoryExportWindow : Window
    {
        public MultiCategoryExportWindow(Document doc)
        {
            InitializeComponent();

            // Khởi tạo ViewModel với Document thực tế của Revit
            this.DataContext = new MultiCategoryExportViewModel(doc);
        }
    }
}