using Serilog;
using System;
using System.Windows;
using System.Windows.Input;

namespace ChipbankImport
{

    public partial class SpecialModal : Window
    {
        public string? getName { get; set; } // from LoginModal
        public string? getID { get; set; } // from LoginModal
        public SpecialModal()
        {
            InitializeComponent();
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                DragMove();
            }
        }

        private void ExitModalSpecial_Click(object sender, RoutedEventArgs e)
        {
            Log.Logger = new LoggerConfiguration()
               .MinimumLevel.Information()
               .WriteTo.File("LogSpecialLOT.txt")
               .CreateLogger();
            try
            {
                Log.Information($"{getID} {getName}");
            }
            catch (Exception ex)
            {
                Log.Error(ex.Message);
            }
            finally
            {
                Log.CloseAndFlush();
            }
            Close();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            labelID.Text = $"{getName}";
        }
    }
}
