using System.Windows;
using System.Windows.Input;

namespace ChipbankImport
{
    public partial class CustomMessageBox : Window
    {
        public CustomMessageBox()
        {
            InitializeComponent();
        }

        private void ExitAlarm_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void AlarmOK_Click(object sender, RoutedEventArgs e)
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
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = true;
            this.Visibility = Visibility.Hidden;
        }
    }
}
