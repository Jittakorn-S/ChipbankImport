using System.Windows;
using System.Windows.Input;

namespace ChipbankImport
{
    /// <summary>
    /// Interaction logic for ModalCondition.xaml
    /// </summary>
    public partial class ModalCondition : Window
    {
        public static bool setIsyes; // To UploadButton_Click ModalFD
        public static bool setisyesSample; // To submitButton_Click Mainwindow
        public ModalCondition()
        {
            InitializeComponent();
        }
        private void selectNo_Click(object sender, RoutedEventArgs e)
        {
            setIsyes = false;
            setisyesSample = false;
            Close();
        }
        private void selectYes_Click(object sender, RoutedEventArgs e)
        {
            setisyesSample = true;
            setIsyes = true;
            Close();
        }
        private void ExitAlarm_Click(object sender, RoutedEventArgs e)
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
