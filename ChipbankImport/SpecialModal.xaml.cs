using Serilog;
using System;
using System.Configuration;
using System.IO;
using System.IO.Compression;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace ChipbankImport
{

    public partial class SpecialModal : Window
    {
        ProgressBar progress = new ProgressBar();
        public string? getName { get; set; } // from LoginModal
        public string? getID { get; set; } // from LoginModal
        int totalFiles = 0;
        int processedFiles = 0;

        public SpecialModal()
        {
            InitializeComponent();
            Log.Logger = new LoggerConfiguration()
               .MinimumLevel.Information()
               .WriteTo.File("LogSpecialLOT.txt")
               .CreateLogger();
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
            Close();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            labelID.Text = $"{getName}";
            barcodeInput.Focus();
        }

        private void barcodeInput_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                enterButton_Click(sender, e);
            }
        }

        private void enterButton_Click(object sender, RoutedEventArgs e)
        {
            int CountText = barcodeInput.Text.Length;
            if (barcodeInput.Text.StartsWith('$') || CountText == 16 || CountText == 14)
            {
                if (CountText == 16)
                {
                    string filename = barcodeInput.Text + ".zip";
                    MainWindow.Unzip(filename);
                }
                else if (barcodeInput.Text.StartsWith('$'))
                {
                    Application.Current.Dispatcher.Invoke((Action)async delegate
                    {
                        string sampleLot = barcodeInput.Text.Trim('$');
                        MainWindow.AlarmConditionBox($"Confirm overwrite lot: {sampleLot} ?");
                        if (ModalCondition.setisyesSample)
                        {
                            progress.Show();
                            await UnzipSampleLot(sampleLot);
                            progress.Close();
                        }
                    });
                }
                else if (CountText == 14)
                {
                    ModalEDSSlip modalEDSSlip = new ModalEDSSlip();
                    modalEDSSlip.zipfileName = barcodeInput.Text;
                    modalEDSSlip.ShowDialog();
                }
                else
                {
                    MainWindow.AlarmBox("Barcode Mismatch !!!");
                    barcodeInput.Focus();
                }
                barcodeInput.Clear();
                barcodeInput.Focus();
            }
            else
            {
                MainWindow.AlarmBox("Barcode Mismatch !!!");
                barcodeInput.Focus();
            }
        }
        private async Task UnzipSampleLot(string getSamplelot)
        {
            string? extractPath = ConfigurationManager.AppSettings["ExtractPath"];
            string? ChecklotName = ConfigurationManager.AppSettings["ChecklotName"]; /*Shared Folder*/
            bool checkLot = false;
            bool checkFolderlot = false;

            DirectoryInfo directoryInfo = new DirectoryInfo(ChecklotName!);
            FileInfo[] files = directoryInfo.GetFiles();

            totalFiles = files.Length;
            processedFiles = 0;

            await Task.Run(() =>
            {
                foreach (FileInfo file in files)
                {
                    string[] splitName = file.Name.Split('.');
                    string lotName = splitName[0];
                    string trimLot = getSamplelot;
                    string lotPathFolder = Path.Combine(extractPath!, trimLot);

                    if (lotName == trimLot)
                    {
                        if (!file.FullName.EndsWith(".bak"))
                        {
                            if (file.Exists)
                            {
                                try
                                {
                                    ZipFile.ExtractToDirectory(file.FullName, lotPathFolder, true);
                                    File.Move(file.FullName, file.FullName + ".bak");
                                    Log.Information($"{getID} {getName} Confirm overwrite lot : {lotName}");
                                    Log.CloseAndFlush();
                                    checkLot = true;
                                    break;
                                }
                                catch (Exception e)
                                {
                                    MainWindow.AlarmBox(e.Message);
                                    break;
                                }
                            }
                            else
                            {
                                Application.Current.Dispatcher.Invoke(() => MainWindow.AlarmBox("Not found LOT file, please check !!!"));
                                break;
                            }
                        }
                        else
                        {
                            Application.Current.Dispatcher.Invoke(() => MainWindow.AlarmBox("Unzip already, please check zip file !!!"));
                            break;
                        }
                    }

                    processedFiles++;

                    Application.Current.Dispatcher.Invoke(async () =>
                    {
                        double progressPercentage = (double)processedFiles / totalFiles * 100;
                        if (progressPercentage != 0 && !double.IsInfinity(progressPercentage))
                        {
                            await Task.Delay((int)progressPercentage);
                        }
                        else
                        {
                            await Task.Delay(1000);
                        }
                    });
                }
            });

            if (checkLot)
            {
                MainWindow.AlarmBox("Upload Successfully !!!");
            }
            if (!checkFolderlot && !checkLot)
            {
                MainWindow.AlarmBox("Not found zip file in CBAll !!!");
            }
        }
    }
}
