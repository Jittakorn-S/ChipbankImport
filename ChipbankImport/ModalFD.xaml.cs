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
    /// Interaction logic for ModalFD.xaml
    /// </summary>
    public partial class ModalFD : Window
    {
        public string? _InvoiceNo { get; set; } // from Mainwindow
        public int _LotCount { get; set; } // from Mainwindow
        public ModalFD()
        {
            InitializeComponent();
        }

        private void exitModal_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void FDModal_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                DragMove();
            }
        }

        private void InvoiceTextbox_Loaded(object sender, RoutedEventArgs e)
        {
            if (_InvoiceNo != null)
            {
                InvoiceTextbox.Text = _InvoiceNo;
            }
            InvoiceTextbox.Text = _InvoiceNo;
            LotNoTextbox.Text = $"{_LotCount} Lots";
        }
    }
}
