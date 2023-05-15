using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using System.Windows.Shapes;

namespace ChipbankImport
{
    /// <summary>
    /// Interaction logic for ModalCondition.xaml
    /// </summary>
    public partial class ModalCondition : Window
    {
        public static bool setIsyes; // To ModalFD
        public ModalCondition()
        {
            InitializeComponent();
        }
        private void selectNo_Click(object sender, RoutedEventArgs e)
        {
            setIsyes = false;
            Close();
        }
        private void selectYes_Click(object sender, RoutedEventArgs e)
        {
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
