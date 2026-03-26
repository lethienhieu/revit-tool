using System.Collections.ObjectModel;
using System.Windows;

namespace THBIM
{
    public partial class CombineOrderDialog : Window
    {
        public ObservableCollection<SheetItem> OrderedList { get; set; }

        public CombineOrderDialog(ObservableCollection<SheetItem> items)
        {
            InitializeComponent();

            // Tạo một bản sao để người dùng sắp xếp. Tránh việc đang làm lỡ tay ấn X tắt cửa sổ thì list gốc bị lỗi.
            OrderedList = new ObservableCollection<SheetItem>(items);
            DgOrder.ItemsSource = OrderedList;

            TxtTotal.Text = $"Total number of items {OrderedList.Count}";
        }

        private void BtnUp_Click(object sender, RoutedEventArgs e)
        {
            int i = DgOrder.SelectedIndex;
            if (i > 0)
            {
                OrderedList.Move(i, i - 1);
                DgOrder.ScrollIntoView(DgOrder.SelectedItem);
            }
        }

        private void BtnDown_Click(object sender, RoutedEventArgs e)
        {
            int i = DgOrder.SelectedIndex;
            if (i >= 0 && i < OrderedList.Count - 1)
            {
                OrderedList.Move(i, i + 1);
                DgOrder.ScrollIntoView(DgOrder.SelectedItem);
            }
        }

        private void BtnTop_Click(object sender, RoutedEventArgs e)
        {
            int i = DgOrder.SelectedIndex;
            if (i > 0)
            {
                OrderedList.Move(i, 0);
                DgOrder.ScrollIntoView(DgOrder.SelectedItem);
            }
        }

        private void BtnBottom_Click(object sender, RoutedEventArgs e)
        {
            int i = DgOrder.SelectedIndex;
            if (i >= 0 && i < OrderedList.Count - 1)
            {
                int last = OrderedList.Count - 1;
                OrderedList.Move(i, last);
                DgOrder.ScrollIntoView(DgOrder.SelectedItem);
            }
        }

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }
    }
}