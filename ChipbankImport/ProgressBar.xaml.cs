using System.Windows;

namespace ChipbankImport
{
    public partial class ProgressBar : Window
    {
        public ProgressBar()
        {
            InitializeComponent();
        }
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = true;
            this.Visibility = Visibility.Hidden;
        }
    }
}
