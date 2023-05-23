using System.Windows;
using System.Windows.Input;

namespace ChipbankImport
{
    public partial class DataWafer : Window
    {
        public DataWafer()
        {
            InitializeComponent();
        }

        private void ExitModalFD_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void closeButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Card_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                DragMove();
            }
        }
    }
}
