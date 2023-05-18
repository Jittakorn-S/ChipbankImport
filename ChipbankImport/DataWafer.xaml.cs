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
using System.Windows.Shapes;

namespace ChipbankImport
{
    /// <summary>
    /// Interaction logic for DataWafer.xaml
    /// </summary>
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
