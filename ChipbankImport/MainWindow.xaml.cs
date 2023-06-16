using System;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace ChipbankImport
{
    public partial class MainWindow : Window
    {
        private static CustomMessageBox? CustomMessageBox;
        private static ModalCondition? ModalCondition;
        public static readonly string? AlarmMessage;
        ProgressBar progress = new ProgressBar();
        int totalFiles = 0;
        int processedFiles = 0;
        public MainWindow()
        {
            InitializeComponent();
            InputLanguageManager.SetInputLanguage(TextInputBarcode, CultureInfo.CreateSpecificCulture("en-US"));
            TextInputBarcode.Focus();
        }
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            restoreScreenButton.Visibility = Visibility.Collapsed;
            WindowState = WindowState.Normal;
        }
        private void TextInputBarcode_Initialized(object sender, EventArgs e)
        {
            InputLanguageManager.SetInputLanguage(TextInputBarcode, CultureInfo.CreateSpecificCulture("en-US"));
        }
        private void TextInputBarcode_GotFocus(object sender, RoutedEventArgs e)
        {
            InputLanguageManager.SetInputLanguage(TextInputBarcode, CultureInfo.CreateSpecificCulture("en-US"));
        }
        private void exitButton_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }
        private void sideExitButton_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }
        private void specialButton_Click(object sender, RoutedEventArgs e)
        {
            LoginModal loginModal = new LoginModal();
            loginModal.Show();
        }

        private void Card_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                DragMove();
            }
        }
        private void minimizeButton_Click(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState.Minimized;
        }
        private void restoreScreenButton_Click(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState.Normal;
            WindowStyle = WindowStyle.None;
            ResizeMode = ResizeMode.CanResize;
        }
        private void fullScreenButton_Click(object sender, RoutedEventArgs e)
        {
            if (WindowState == WindowState.Normal)
            {
                WindowState = WindowState.Maximized;
                WindowStyle = WindowStyle.None;
                ResizeMode = ResizeMode.NoResize;
            }
        }
        private void TextBox_KeyDown(object sender, KeyEventArgs e)
        {
            InputLanguageManager.SetInputLanguage(TextInputBarcode, CultureInfo.CreateSpecificCulture("en-US"));
            if (e.Key == Key.Enter)
            {
                submitButton_Click(sender, e);
            }
        }
        private void Window_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2 || WindowState == WindowState.Normal)
            {
                WindowState = WindowState.Maximized;
                WindowStyle = WindowStyle.None;
                ResizeMode = ResizeMode.NoResize;
            }
            else
            {
                WindowState = WindowState.Normal;
                WindowStyle = WindowStyle.None;
                ResizeMode = ResizeMode.CanResize;
            }
        }
        private void Window_StateChanged(object sender, EventArgs e)
        {
            switch (WindowState)
            {
                case WindowState.Maximized:
                    fullScreenButton.Visibility = Visibility.Collapsed;
                    restoreScreenButton.Visibility = Visibility.Visible;
                    break;
                case WindowState.Normal:
                    fullScreenButton.Visibility = Visibility.Visible;
                    restoreScreenButton.Visibility = Visibility.Collapsed;
                    break;
                default:
                    restoreScreenButton.Visibility = Visibility.Collapsed;
                    break;
            }
        }
        private void submitButton_Click(object sender, RoutedEventArgs e)
        {
            int CountText = TextInputBarcode.Text.Length;
            if (TextInputBarcode.Text.StartsWith('$') || CountText == 16 || CountText == 14)
            {
                if (CountText == 16)
                {
                    string filename = TextInputBarcode.Text + ".zip";
                    Unzip(filename);
                }
                else if (TextInputBarcode.Text.StartsWith('$'))
                {
                    Application.Current.Dispatcher.Invoke((Action)async delegate
                    {
                        string sampleLot = TextInputBarcode.Text.Trim('$');
                        AlarmConditionBox($"Confirm Upload Sample Lot : {sampleLot} ?");
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
                    modalEDSSlip.zipfileName = TextInputBarcode.Text;
                    modalEDSSlip.ShowDialog();
                }
                else
                {
                    AlarmBox("Barcode Mismatch !!!");
                    TextInputBarcode.Clear();
                    TextInputBarcode.Focus();
                }
                TextInputBarcode.Clear();
                TextInputBarcode.Focus();
            }
            else
            {
                AlarmBox("Barcode Mismatch !!!");
                TextInputBarcode.Clear();
                TextInputBarcode.Focus();
            }
        }
        public static void Unzip(string FileName)
        {
            string? FileToCopy = ConfigurationManager.AppSettings["FileToCopyPath"] + FileName; /*Shared Folder*/
            string? NewCopyCB = ConfigurationManager.AppSettings["NewCopyCBPath"]!; /*for backup file before sending*/
            string? ProcessPath = ConfigurationManager.AppSettings["ProcessPath"]!;
            string? ChangeDirectory = ConfigurationManager.AppSettings["ChangeDirectory"]!;
            if (Directory.Exists(ProcessPath))
            {
                Directory.Delete(ProcessPath, true);
            }

            Directory.CreateDirectory(ProcessPath);

            if (System.IO.File.Exists(NewCopyCB))
            {
                System.IO.File.Delete(NewCopyCB);
            }
            if (System.IO.File.Exists(FileToCopy))
            {
                try
                {
                    System.IO.File.Copy(FileToCopy, NewCopyCB);
                }
                catch (Exception)
                {
                    AlarmBox($"Please check location file {FileToCopy} or {NewCopyCB}");
                }
            }
            try
            {
                Directory.SetCurrentDirectory(ChangeDirectory);
                ZipFile.ExtractToDirectory(NewCopyCB, ProcessPath);
            }
            catch (Exception)
            {
                AlarmBox($"Please check location file {NewCopyCB} or {ProcessPath}");
            }

            if (System.IO.File.Exists(ProcessPath + @"\FD\Refidc02.fd"))
            {
                string? ReadlineText = null;
                string? InvoiceNo = null;
                int LotCount = 0;
                using (FileStream fileStream = new FileStream(ProcessPath + @"\FD\Refidc02.fd", FileMode.Open))
                {
                    using (StreamReader streamReader = new StreamReader(fileStream))
                    {
                        while ((ReadlineText = streamReader.ReadLine()) != null)
                        {
                            InvoiceNo = ReadlineText.Substring(86, 8);
                            LotCount = LotCount + 1;
                        }
                    }
                }
                ModalFD Modalfd = new ModalFD();
                Modalfd._InvoiceNo = InvoiceNo!;
                Modalfd._LotCount = LotCount;
                Modalfd.GetProcessPath = ProcessPath + @"\FD\Refidc02.fd";
                Modalfd.ShowDialog();
            }
            else
            {
                AlarmBox("Data not exist check Refidc02.fd !!!");
            }
        }
        private async Task UnzipSampleLot(string getSamplelot)
        {
            string? extractPath = ConfigurationManager.AppSettings["ExtractPath"];
            string? ChecklotName = ConfigurationManager.AppSettings["ChecklotName"]; /*Shared Folder*/
            bool checkLot = false;
            bool checkFolderlot = false;

            DirectoryInfo directoryInfo = new DirectoryInfo(ChecklotName!);
            DirectoryInfo extractPathcheck = new DirectoryInfo(extractPath!);

            DirectoryInfo[] folder = extractPathcheck.GetDirectories();
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
                                for (int i = 0; i < folder.Length; i++)
                                {
                                    DirectoryInfo checkfolder = folder[i];
                                    string folderName = checkfolder.Name;
                                    if (folderName == lotName)
                                    {
                                        checkFolderlot = true;
                                    }
                                }
                                if (!checkFolderlot)
                                {
                                    try
                                    {
                                        ZipFile.ExtractToDirectory(file.FullName, lotPathFolder);
                                        File.Move(file.FullName, file.FullName + ".bak");
                                        checkLot = true;
                                    }
                                    catch
                                    {
                                        break;
                                    }
                                }
                                else
                                {
                                    Application.Current.Dispatcher.Invoke(() => AlarmBox("Unzip already, please check !!!"));
                                    break;
                                }
                            }
                            else
                            {
                                Application.Current.Dispatcher.Invoke(() => AlarmBox("Not found LOT file, please check !!!"));
                            }
                        }
                        else
                        {
                            Application.Current.Dispatcher.Invoke(() => AlarmBox("Unzip already, please check !!!"));
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
                AlarmBox("Upload Successfully !!!");
            }
            if (!checkFolderlot && !checkLot)
            {
                AlarmBox("Not found zip file in CBAll !!!");
            }
        }
        public static void AlarmBox(string AlarmMessage)
        {
            CustomMessageBox = new CustomMessageBox();
            CustomMessageBox.AlarmText.Content = AlarmMessage;
            CustomMessageBox.ShowDialog();
        }
        public static void AlarmConditionBox(string? textMessage)
        {
            ModalCondition = new ModalCondition();
            ModalCondition.textContent.Content = textMessage;
            ModalCondition.ShowDialog();
        }
    }
}
