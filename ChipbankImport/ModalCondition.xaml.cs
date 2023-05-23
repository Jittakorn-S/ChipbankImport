using System.Windows;
using System.Windows.Input;

namespace ChipbankImport
{
    public partial class ModalCondition : Window
    {
        public static bool setIsyes { get; set; } // To UploadButton_Click ModalFD
        public static bool setisyesSample { get; set; } // To submitButton_Click Mainwindow
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
