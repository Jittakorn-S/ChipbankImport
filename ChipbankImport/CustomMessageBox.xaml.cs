using System.Windows;
using System.Windows.Input;

namespace ChipbankImport
{
    /// <summary>
    /// Interaction logic for CustomMessageBox.xaml
    /// </summary>
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
    }
}
