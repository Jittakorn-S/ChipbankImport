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
using static ChipbankImport.MainWindow;
using static System.Net.WebRequestMethods;
using File = System.IO.File;

namespace ChipbankImport
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private static CustomMessageBox? CustomMessageBox;
        private static string? TmpData;
        private static string? useqno;
        private static string? finseqno;

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
                    }
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
        public struct WaferData
        {
            public string ActualNo;
            public string WFLotNo;
            public string RFSeqNo;
            public string ChipModelName;
            public string ModelCode1;
            public string ModelCode2;
            public string RohmModelName;
            public string InvoiceNo;
            public string CaseNo;
            public string Box;
            public string OutDiv;
            public string RecDiv;
            public string OrderNo;
            public string ControlCode;
            public string PayClass;
            public string WFCount;
            public string ChipCount;
            public string WaferDATA;
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
                            WaferData waferData = new WaferData
                            {
                                ActualNo = ReadlineTextFD.Substring(0, 10),
                                WFLotNo = ReadlineTextFD.Substring(10, 11),
                                RFSeqNo = ReadlineTextFD.Substring(22, 10),
                                ChipModelName = ReadlineTextFD.Substring(32, 20),
                                ModelCode1 = ReadlineTextFD.Substring(52, 8),
                                ModelCode2 = ReadlineTextFD.Substring(60, 6),
                                RohmModelName = ReadlineTextFD.Substring(66, 14),
                                InvoiceNo = ReadlineTextFD.Substring(86, 7),
                                CaseNo = ReadlineTextFD.Substring(94, 6),
                                Box = ReadlineTextFD.Substring(100, 7),
                                OutDiv = ReadlineTextFD.Substring(110, 5),
                                RecDiv = ReadlineTextFD.Substring(115, 5),
                                OrderNo = ReadlineTextFD.Substring(120, 12),
                                ControlCode = ReadlineTextFD.Substring(135, 7),
                                PayClass = ReadlineTextFD.Substring(144, 1),
                                WFCount = ReadlineTextFD.Substring(143, 2),
                                ChipCount = ReadlineTextFD.Substring(147, 5),
                                WaferDATA = ReadlineTextFD.Substring(154, 358)
                            };

                            SET_WAFER(ReadlineTextFD);
                            SETSEQ();
                            STOCKINDATA(waferData);
                            STOCKDATA(waferData);
                        }
                    }
                }
            }
            else
            {
                AlarmBox("Data not exist check Refidc02.fd !!!");
            }
        }
        public static void SET_WAFER(string GetReadlineTextFD)
        {
            int POS = 152;
            int POS2 = 155;
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
                    break;
                }
            }
        }
        public static void STOCKINDATA(WaferData GetwaferData)
        {
            string TIMESTAMP = DateTime.Now.ToString();
            string STOCKDATE = DateTime.Now.ToString("yyMMdd");
            string WFDATA1 = TmpData!.Substring(0, 180);
            string WFDATA2 = TmpData!.Substring(180, 180);
            string ReWFDATA1 = WFDATA1.Replace("  0", "   ");
            string ReWFDATA2 = WFDATA2.Replace("  0", "   ");
            string ConnetionString = ConfigurationManager.AppSettings["ConnetionString"]!;
            string sqlInsert = "INSERT INTO CHIPNYUKO (CHIPMODELNAME, MODELCODE1, MODELCODE2, WFLOTNO, SEQNO, OUTDIV," +
                               " RECDIV, STOCKDATE, WFCOUNT, CHIPCOUNT, SLIPNO, SLIPNOEDA, ORDERNO, INVOICENO, HOLDFLAG, CASENO, WFDATA1, WFDATA2, WFINPUT, " +
                               "TIMESTAMP, RFSEQNO) " +
                               "VALUES (@CHIPMODELNAME, @MODELCODE1, @MODELCODE2, @WFLOTNO, @SEQNO, @OUTDIV, @RECDIV, @STOCKDATE, @WFCOUNT, " +
                               "@CHIPCOUNT, @SLIPNO, @SLIPNOEDA, @ORDERNO, @INVOICENo, @HOLDFLAG, @CASENO, @tmpdata1, @tmpdata2, @WFINPUT, @TIMESTAMP, @RF_SEQNO)";
            string sqlUpdate = "UPDATE CHIPNYUKO SET PLASMA = (SELECT PLASMA FROM CHIPMASTER Where CHIPMODELNAME = @CHIPMODELNAME AND PLASMA = '1') " +
                               "WHERE WFLOTNO = @WFLOTNO AND SEQNO = @finseqno";
            using (SqlConnection connection = new SqlConnection(ConnetionString))
            {
                try
                {
                    connection.Open();
                    SqlCommand sqlCommandQuery = new SqlCommand(sqlInsert, connection);
                    sqlCommandQuery.Parameters.AddWithValue("@CHIPMODELNAME", GetwaferData.ChipModelName);
                    sqlCommandQuery.Parameters.AddWithValue("@MODELCODE1", GetwaferData.ModelCode1);
                    sqlCommandQuery.Parameters.AddWithValue("@MODELCODE2", GetwaferData.ModelCode2);
                    sqlCommandQuery.Parameters.AddWithValue("@WFLOTNO", GetwaferData.WFLotNo);
                    sqlCommandQuery.Parameters.AddWithValue("@SEQNO", finseqno);
                    sqlCommandQuery.Parameters.AddWithValue("@OUTDIV", GetwaferData.OutDiv);
                    sqlCommandQuery.Parameters.AddWithValue("@RECDIV", GetwaferData.RecDiv);
                    sqlCommandQuery.Parameters.AddWithValue("@STOCKDATE", STOCKDATE);
                    sqlCommandQuery.Parameters.AddWithValue("@WFCOUNT", GetwaferData.WFCount);
                    sqlCommandQuery.Parameters.AddWithValue("@CHIPCOUNT", GetwaferData.ChipCount);
                    sqlCommandQuery.Parameters.AddWithValue("@SLIPNO", "          ");
                    sqlCommandQuery.Parameters.AddWithValue("@SLIPNOEDA", "  ");
                    sqlCommandQuery.Parameters.AddWithValue("@ORDERNO", GetwaferData.OrderNo);
                    sqlCommandQuery.Parameters.AddWithValue("@INVOICENo", GetwaferData.InvoiceNo);
                    sqlCommandQuery.Parameters.AddWithValue("@HOLDFLAG", "");
                    sqlCommandQuery.Parameters.AddWithValue("@CASENO", GetwaferData.CaseNo);
                    sqlCommandQuery.Parameters.AddWithValue("@tmpdata1", ReWFDATA1);
                    sqlCommandQuery.Parameters.AddWithValue("@tmpdata2", ReWFDATA2);
                    sqlCommandQuery.Parameters.AddWithValue("@WFINPUT", "1");
                    sqlCommandQuery.Parameters.AddWithValue("@TIMESTAMP", TIMESTAMP);
                    sqlCommandQuery.Parameters.AddWithValue("@RF_SEQNO", GetwaferData.RFSeqNo);
                    //sqlCommandQuery.ExecuteNonQuery();
                }
                catch (SqlException)
                {
                    AlarmBox("Can not insert data or connect to the database !!!");
                }
                try
                {
                    SqlCommand sqlCommandQuery = new SqlCommand(sqlUpdate, connection);
                    sqlCommandQuery.Parameters.AddWithValue("@CHIPMODELNAME", GetwaferData.ChipModelName);
                    sqlCommandQuery.Parameters.AddWithValue("@WFLOTNO", GetwaferData.WFLotNo);
                    sqlCommandQuery.Parameters.AddWithValue("@finseqno", finseqno);
                    sqlCommandQuery.ExecuteNonQuery();
                }
                catch (SqlException)
                {
                    AlarmBox("Can not update data or connect to the database !!!");
                }
            }
        }
        public static void SETSEQ()
        {
            try
            {
                string? ALOCATEDATE = null;
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
                                    AlarmBox("Can not connect to the database !!!");
                                }
                                finally
                                {
                                    connection.Close();
                                }
                            }
                            finseqno = $"Q{ALOCATEDATE!.Substring(1, 5)}{useqno}";
                        }
                    }
                }
            }
            catch (Exception e)
            {
                AlarmBox(e.Message);
            }
        }
        public static void STOCKDATA(WaferData GetwaferData)
        {
            try
            {
                string STOCKDATE = DateTime.Now.ToString("yyMMdd");
                string TIMESTAMP = DateTime.Now.ToString();
                string ConnectionString = ConfigurationManager.AppSettings["ConnetionString"]!;
                string sqlInsert = "INSERT INTO CHIPZAIKO (CHIPMODELNAME, MODELCODE1, MODELCODE2, WFLOTNO, SEQNO, ENO, LOCATION, WFCOUNT, CHIPCOUNT, STOCKDATE, " +
                                   "RETURNFLAG, REMAINFLAG, HOLDFLAG, STAFFNO, PREOUTFLAG, INVOICENO, PROCESSCODE, DELETEFLAG, TIMESTAMP)" +
                                   "VALUES (@CHIPMODELNAME, @MODELCODE1, @MODELCODE2, @WFLOTNO, @SEQNO, @ENO, @LOCATION, @WFCOUNT, @CHIPCOUNT, @STOCKDATE, " +
                                   "@RETURNFLAG, @REMAINFLAG, @HOLDFLAG, @STAFFNO, @PREOUTFLAG, @INVOICENO, @PROCESSCODE, @DELETEFLAG, @TIMESTAMP)";
                using (SqlConnection connection = new SqlConnection(ConnectionString))
                {
                    connection.Open();
                    SqlCommand sqlCommandQuery = new SqlCommand(sqlInsert, connection);
                    sqlCommandQuery.Parameters.AddWithValue("@CHIPMODELNAME", GetwaferData.ChipModelName);
                    sqlCommandQuery.Parameters.AddWithValue("@MODELCODE1", GetwaferData.ModelCode1);
                    sqlCommandQuery.Parameters.AddWithValue("@MODELCODE2", GetwaferData.ModelCode2);
                    sqlCommandQuery.Parameters.AddWithValue("@WFLOTNO", GetwaferData.WFLotNo);
                    sqlCommandQuery.Parameters.AddWithValue("@SEQNO", finseqno);
                    sqlCommandQuery.Parameters.AddWithValue("@ENO", "");
                    sqlCommandQuery.Parameters.AddWithValue("@LOCATION", "");
                    sqlCommandQuery.Parameters.AddWithValue("@WFCOUNT", GetwaferData.WFCount);
                    sqlCommandQuery.Parameters.AddWithValue("@CHIPCOUNT", GetwaferData.ChipCount);
                    sqlCommandQuery.Parameters.AddWithValue("@STOCKDATE", STOCKDATE);
                    sqlCommandQuery.Parameters.AddWithValue("@RETURNFLAG", "");
                    sqlCommandQuery.Parameters.AddWithValue("@REMAINFLAG", "");
                    sqlCommandQuery.Parameters.AddWithValue("@HOLDFLAG", "");
                    sqlCommandQuery.Parameters.AddWithValue("@STAFFNO", "00001");
                    sqlCommandQuery.Parameters.AddWithValue("@PREOUTFLAG", "");
                    sqlCommandQuery.Parameters.AddWithValue("@INVOICENO", GetwaferData.InvoiceNo);
                    sqlCommandQuery.Parameters.AddWithValue("@PROCESSCODE", GetwaferData.RecDiv);
                    sqlCommandQuery.Parameters.AddWithValue("@DELETEFLAG", "");
                    sqlCommandQuery.Parameters.AddWithValue("@TIMESTAMP", TIMESTAMP);
                    sqlCommandQuery.ExecuteNonQuery();
                }
            }
            catch (SqlException)
            {
                AlarmBox("Can not insert data or connect to the database !!!");
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
