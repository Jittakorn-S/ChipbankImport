using System;
using System.Configuration;
using System.IO;
using System.IO.Compression;
using System.Windows;
using System.Windows.Input;

namespace ChipbankImport
{
    public partial class MainWindow : Window
    {
        private static CustomMessageBox? CustomMessageBox;
        private static ModalCondition? ModalCondition;
        internal static string? AlarmMessage;

        public MainWindow()
        {
            InitializeComponent();
            TextInputBarcode.Focus();
        }
        private void exitButton_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
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
            Window window = Application.Current.MainWindow;
            window.WindowState = WindowState.Minimized;
        }
        private void fullScreenButton_Click(object sender, RoutedEventArgs e)
        {
            Window window = Application.Current.MainWindow;
            if (window.WindowState == WindowState.Maximized)
            {
                window.WindowState = WindowState.Normal;
            }
            else
            {
                window.WindowState = WindowState.Maximized;
            }
        }
        private void TextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                submitButton_Click(sender, e);
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
                    string sampleLot = TextInputBarcode.Text.Trim('$');
                    AlarmConditionBox($"Confirm Upload Sample Lot : {sampleLot} ?");
                    if (ModalCondition.setisyesSample == true)
                    {
                        UnzipSampleLot(sampleLot);
                    }
                }
                else if (CountText == 14)
                {
                    ModalSampleLot modalSampleLot = new ModalSampleLot();
                    modalSampleLot.zipfileName = TextInputBarcode.Text;
                    modalSampleLot.ShowDialog();
                }
                else
                {
                    AlarmBox("Barcode Mismatch !!!");
                    TextInputBarcode.Focus();
                }
                TextInputBarcode.Clear();
                TextInputBarcode.Focus();
            }
            else
            {
                AlarmBox("Barcode Mismatch !!!");
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

            if (File.Exists(NewCopyCB))
            {
                File.Delete(NewCopyCB);
            }
            if (File.Exists(FileToCopy))
            {
                try
                {
                    File.Copy(FileToCopy, NewCopyCB);
                }
                catch (Exception)
                {
                    AlarmBox("Please check location file");
                }
            }
            try
            {
                Directory.SetCurrentDirectory(ChangeDirectory);
                ZipFile.ExtractToDirectory(NewCopyCB, ProcessPath);
            }
            catch (Exception)
            {
                AlarmBox("Please check location file");
            }

            if (File.Exists(ProcessPath + @"\FD\Refidc02.fd"))
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
        private void UnzipSampleLot(string getSamplelot)
        {
            string? extractPath = ConfigurationManager.AppSettings["ExtractPath"];
            string? ChecklotName = ConfigurationManager.AppSettings["ChecklotName"]; /*Shared Folder*/
            bool checkLot = false;
            DirectoryInfo directoryInfo = new DirectoryInfo(ChecklotName!);
            FileInfo[] files = directoryInfo.GetFiles();
            foreach (FileInfo file in files)
            {

                string[] splitName = file.Name.Split('.');
                string lotName = splitName[0];
                string trimLot = getSamplelot;
                string LotpathFolder = extractPath! + trimLot;
                if (lotName == trimLot)
                {
                    if (Directory.Exists(LotpathFolder))
                    {
                        Directory.Delete(LotpathFolder, true);
                        ZipFile.ExtractToDirectory(file.FullName, LotpathFolder);
                        checkLot = true;
                    }
                    else
                    {
                        ZipFile.ExtractToDirectory(file.FullName, LotpathFolder);
                        checkLot = true;
                    }
                }
            }
            if (checkLot == true)
            {
                AlarmBox("Upload Successfully !!!");
            }
            else
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
