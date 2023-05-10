using MaterialDesignThemes.Wpf;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Formats.Asn1;
using System.IO;
using System.IO.Compression;
using System.IO.Pipes;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Shell;
using static System.Net.WebRequestMethods;
using File = System.IO.File;

namespace ChipbankImport
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        static CustomMessageBox? CustomMessageBox;
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
            int CountText = TextInputBarcode.Text.Count();
            if (CountText == 16)
            {
                string filename = TextInputBarcode.Text + ".zip";
                Unzip(filename);
            }
            else
            {
                AlarmBox("Please enter the data correctly");
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
                            InvoiceNo = ReadlineText.Substring(86, 7);
                            LotCount = LotCount + 1;
                        }
                        streamReader.Close();
                        streamReader.Dispose();
                    }
                    fileStream.Close();
                    fileStream.Dispose();
                }
                DataFDSheet(ProcessPath + @"\FD\Refidc02.fd"); /*Prepare Data*/
                ModalFD Modalfd = new ModalFD();
                Modalfd._InvoiceNo = InvoiceNo!;
                Modalfd._LotCount = LotCount;
                Modalfd.ShowDialog();
            }
            else
            {
                AlarmBox("Data not exist check Refidc02.fd !!!");
            }

        }
        public static void DataFDSheet(string GetProcessPath)
        {
            string? ReadlineTextFD = null;
            if (File.Exists(GetProcessPath))
            {
                using (FileStream fileStream = new FileStream(GetProcessPath, FileMode.Open))
                {
                    using (StreamReader streamReader = new StreamReader(fileStream))
                    {
                        while ((ReadlineTextFD = streamReader.ReadLine()) != null)
                        {
                            string ACTULNO = ReadlineTextFD.Substring(0, 10);
                            string WFLOTNO = ReadlineTextFD.Substring(10, 11);
                            string RFSEQNO = ReadlineTextFD.Substring(22, 10);
                            string CHIPMODELNAME = ReadlineTextFD.Substring(32, 13);
                            string MODELCODE1 = ReadlineTextFD.Substring(52, 8);
                            string MODELCODE2 = ReadlineTextFD.Substring(60, 6);
                            string ROHMMODELNAME = ReadlineTextFD.Substring(66, 14);
                            string INVOICENo = ReadlineTextFD.Substring(86, 7);
                            string CASENO = ReadlineTextFD.Substring(94, 6);
                            string BOX = ReadlineTextFD.Substring(100, 7);
                            string OUTDIV = ReadlineTextFD.Substring(110, 5);
                            string RECDIV = ReadlineTextFD.Substring(115, 5);
                            string ORDERNO = ReadlineTextFD.Substring(120, 12);
                            string CONTROLCODE = ReadlineTextFD.Substring(135, 7);
                            string PAYCLASS = ReadlineTextFD.Substring(144, 1);
                            string WFCOUNT = ReadlineTextFD.Substring(143, 2);
                            string CHIPCOUNT = ReadlineTextFD.Substring(147, 5);
                            string WAFERDATA = ReadlineTextFD.Substring(154, 358);

                            SET_WAFER(ReadlineTextFD);
                            SETSEQ();
                        }
                        streamReader.Close();
                        streamReader.Dispose();
                    }
                    fileStream.Close();
                    fileStream.Dispose();
                }
            }

        }
        public static void SET_WAFER(string GetReadlineTextFD)
        {
            int POS = 152;
            int POS2 = 155;
            string? TmpData = null;
            for (int i = 0; i <= 39; i++)
            {
                if (GetReadlineTextFD.Substring(POS + (9 * i), 3).Length != 0)
                {
                    string WFSEQ = GetReadlineTextFD.Substring(POS + (9 * i), 3);
                    string waferchipcount = GetReadlineTextFD.Substring(POS2 + (9 * i), 6);
                    //SET_WF_SEQ
                    if (GetReadlineTextFD.Substring(POS + (9 * i), 3).Length == 1)
                    {
                        TmpData = TmpData + "  " + WFSEQ;
                    }
                    else if (GetReadlineTextFD.Substring(POS + (9 * i), 3).Length == 2)
                    {
                        TmpData = TmpData + " " + WFSEQ;
                    }
                    else if (GetReadlineTextFD.Substring(POS + (9 * i), 3).Length == 3)
                    {
                        TmpData = TmpData + WFSEQ;
                    }
                    //-------------------------------------------------------------
                    //SET_CHIP_COUNT
                    if (waferchipcount.Trim().Length == 1)
                    {
                        TmpData = TmpData + "     " + waferchipcount.Trim();
                    }
                    else if (waferchipcount.Trim().Length == 2)
                    {
                        TmpData = TmpData + "    " + waferchipcount.Trim();
                    }
                    else if (waferchipcount.Trim().Length == 3)
                    {
                        TmpData = TmpData + "   " + waferchipcount.Trim();
                    }
                    else if (waferchipcount.Trim().Length == 4)
                    {
                        TmpData = TmpData + "  " + waferchipcount.Trim();
                    }
                    else if (waferchipcount.Trim().Length == 5)
                    {
                        TmpData = TmpData + " " + waferchipcount.Trim();
                    }
                    else if (waferchipcount.Trim().Length == 6)
                    {
                        TmpData = TmpData + waferchipcount.Trim();
                    }
                    //-------------------------------------------------------------
                }
                else
                {
                    AlarmBox("Can not read data check Refidc02.fd !!!");
                }
            }
        }
        public static void STOCKINDATA(string GetTmpData)
        {
            string TIMESTAMP = DateTime.Now.ToString();
            string? tmpdata1 = null, tmpdata2 = null;
            tmpdata1 = GetTmpData.Substring(0, 180);
            tmpdata2 = GetTmpData.Substring(181, 180);

            //    sql = "INSERT CHIPNYUKO(CHIPMODELNAME,MODELCODE1,MODELCODE2,WFLOTNO,SEQNO,"
            //sql = sql & "OUTDIV,RECDIV,STOCKDATE,WFCOUNT,CHIPCOUNT,ORDERNO,INVOICENO,CASENO,"
            //sql = sql & "WFDATA1,WFDATA2,WFINPUT,TIMESTAMP,RFSEQNO) VALUES('"
            //sql = sql & CHIPMODELNAME & "','"
            //sql = sql & MODELCODE1 & "','"
            //sql = sql & MODELCODE2 & "','"
            //sql = sql & WFLOTNO & "','"
            //sql = sql & finseqno & "','"
            //sql = sql & OUTDIV & "','"
            //sql = sql & RECDIV & "','"
            //sql = sql & Now.ToString("yyMMdd") & "',"
            //sql = sql & WFCOUNT & ","
            //sql = sql & CHIPCOUNT & ",'"
            //sql = sql & ORDERNO & "','"
            //sql = sql & INVOICENo & "','"
            //sql = sql & CASENO & "','"
            //sql = sql & tmpdata1 & "','" & tmpdata2 & "','1','" & TIMESTAMP & "','" & RF_SEQNO & "');"


            //    Dim myCmd As New SqlCommand(sql, CN1)

            //myCmd.Connection.Open()
            //myCmd.ExecuteNonQuery()
            //myCmd.Connection.Close()

            //sql = "UPDATE CHIPNYUKO SET PLASMA=(SELECT PLASMA FROM CHIPMASTER " & _
            //            "Where CHIPMODELNAME='" & CHIPMODELNAME & "' AND PLASMA='1') " & _
            //        "WHERE WFLOTNO='" & WFLOTNO & "' AND SEQNO='" & finseqno & "'"

            //Dim myCmd2 As New SqlCommand(sql, CN1)

            //myCmd2.Connection.Open()
            //myCmd2.ExecuteNonQuery()
            //myCmd2.Connection.Close()
        }
        public static void SETSEQ()
        {
            try
            {
                string? ALOCATEDATE = null;
                string? useqno = null;
                int countSeq = 0;
                int SEQNO = 0;
                string ConnectionString = ConfigurationManager.AppSettings["ConnetionString"]!;
                using (SqlConnection connection = new SqlConnection(ConnectionString))
                {
                    connection.Open();
                    using (SqlCommand sqlCommand = new SqlCommand("SELECT * FROM CHIPSYS WHERE SYSKEY = @SYSKEY", connection))
                    {
                        sqlCommand.Parameters.AddWithValue("@SYSKEY", "01");
                        using (SqlDataReader reader = sqlCommand.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                foreach (DbDataRecord record in reader)
                                {
                                    SEQNO = Convert.ToInt16(record.GetDecimal(2));
                                    ALOCATEDATE = record.GetString(1);
                                    if (ALOCATEDATE != DateTime.Now.ToString("yyMMdd"))
                                    {
                                        ALOCATEDATE = DateTime.Now.ToString("yyMMdd");
                                        using (var sqlCommandUpdate = new SqlCommand("UPDATE CHIPSYS " +
                                                                                     "SET ALOCATEDATE = @ALOCATEDATE, SEQNO = 2 " +
                                                                                     "WHERE SYSKEY = '01';", connection))
                                        {
                                            sqlCommandUpdate.Parameters.AddWithValue("@ALOCATEDATE", ALOCATEDATE);
                                            sqlCommandUpdate.ExecuteNonQuery();
                                            useqno = "0001";
                                        }
                                    }
                                    else
                                    {
                                        countSeq = SEQNO.ToString().Length;
                                        if (countSeq == 4)
                                        {
                                            useqno = SEQNO.ToString();
                                        }
                                        else if (countSeq == 3)
                                        {
                                            useqno = "0" + SEQNO.ToString();
                                        }
                                        else if (countSeq == 2)
                                        {
                                            useqno = "00" + SEQNO.ToString();
                                        }
                                        else if (countSeq == 1)
                                        {
                                            useqno = "000" + SEQNO.ToString();
                                        }
                                        SEQNO = SEQNO + 1;

                                        using (SqlCommand sqlCommandUpdate = new SqlCommand("UPDATE CHIPSYS " +
                                                     "SET SEQNO = @SEQNO " +
                                                     "WHERE SYSKEY = '01';", connection))
                                        {
                                            sqlCommandUpdate.Parameters.AddWithValue("@SEQNO", SEQNO);
                                            sqlCommandUpdate.ExecuteNonQuery();
                                        }
                                    }
                                }
                            }
                            else
                            {
                                string sqlCommandInsert = "INSERT INTO CHIPSYS VALUES(@SYSKEY, @ALOCATEDATE, @SEQNO)";
                                try
                                {
                                    ALOCATEDATE = DateTime.Now.ToString("yyMMdd");
                                    SqlCommand sqlCommandQuery = new SqlCommand(sqlCommandInsert, connection);
                                    sqlCommandQuery.Parameters.AddWithValue("@SYSKEY", "01");
                                    sqlCommandQuery.Parameters.AddWithValue("@ALOCATEDATE", ALOCATEDATE);
                                    sqlCommandQuery.Parameters.AddWithValue("@SEQNO", 2);
                                    sqlCommandQuery.ExecuteNonQuery();
                                    useqno = "0001";
                                }
                                catch (SqlException)
                                {
                                    AlarmBox("Can not connect to database !!!");
                                }
                                finally
                                {
                                    connection.Close();
                                }
                            }
                            string finseqno = $"Q{ALOCATEDATE}{useqno}";
                        }
                    }
                }
            }
            catch (Exception e)
            {
                AlarmBox(e.Message);
            }
        }
        public static void AlarmBox(string AlarmMessage)
        {
            CustomMessageBox = new CustomMessageBox();
            CustomMessageBox.AlarmText.Content = AlarmMessage;
            CustomMessageBox.ShowDialog();
        }
    }
}
