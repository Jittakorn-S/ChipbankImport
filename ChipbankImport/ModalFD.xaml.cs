using System;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Windows;
using System.Windows.Input;

namespace ChipbankImport
{
    public partial class ModalFD : Window
    {
        public int _LotCount { get; set; } // from Mainwindow
        public string? _InvoiceNo { get; set; } // from Mainwindow
        public string? GetProcessPath { get; set; } // from Mainwindow
        private static string? TmpData = null;
        private static string? useqno;
        private static string? finseqno;
        private static string? resultTmpData;
        private bool checkNoInvoiceRows = false;
        public struct WaferData
        {
            public string ActualNo { get; set; }
            public string WFLotNo { get; set; }
            public string RFSeqNo { get; set; }
            public string ChipModelName { get; set; }
            public string ModelCode1 { get; set; }
            public string ModelCode2 { get; set; }
            public string RohmModelName { get; set; }
            public string InvoiceNo { get; set; }
            public string CaseNo { get; set; }
            public string Box { get; set; }
            public string OutDiv { get; set; }
            public string RecDiv { get; set; }
            public string OrderNo { get; set; }
            public string ControlCode { get; set; }
            public string PayClass { get; set; }
            public string WFCount { get; set; }
            public string ChipCount { get; set; }
        }
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
                if (_InvoiceNo == "        ")
                {
                    MainWindow.AlarmBox("Not have an invoice, Please check !!!");
                    InvoiceTextbox.Text = "No Invoice";
                    LotNoTextbox.Text = $"{_LotCount} Lot";
                }
                else
                {
                    InvoiceTextbox.Text = _InvoiceNo;
                    LotNoTextbox.Text = $"{_LotCount} Lot";
                }
            }
            else
            {
                MainWindow.AlarmBox("Data not exist check Refidc02.fd !!!");
            }
        }
        private void UploadDataFDSheet()
        {
            bool checkChipzaiko = false;
            bool checkChipnyoko = false;
            try
            {
                string? ReadlineTextFD = null;
                if (File.Exists(GetProcessPath))
                {
                    using (FileStream fileStream = new(GetProcessPath, FileMode.Open))
                    {
                        using (StreamReader streamReader = new StreamReader(fileStream))
                        {
                            while ((ReadlineTextFD = streamReader.ReadLine()) != null)
                            {
                                WaferData waferData = new WaferData
                                {
                                    ActualNo = ReadlineTextFD.Substring(0, 10),
                                    WFLotNo = ReadlineTextFD.Substring(10, 12),
                                    RFSeqNo = ReadlineTextFD.Substring(22, 10),
                                    ChipModelName = ReadlineTextFD.Substring(32, 20),
                                    ModelCode1 = ReadlineTextFD.Substring(52, 8),
                                    ModelCode2 = ReadlineTextFD.Substring(60, 6),
                                    RohmModelName = ReadlineTextFD.Substring(66, 20),
                                    InvoiceNo = ReadlineTextFD.Substring(86, 8),
                                    CaseNo = ReadlineTextFD.Substring(94, 6),
                                    Box = ReadlineTextFD.Substring(100, 10),
                                    OutDiv = ReadlineTextFD.Substring(110, 5),
                                    RecDiv = ReadlineTextFD.Substring(115, 5),
                                    OrderNo = ReadlineTextFD.Substring(120, 10),
                                    ControlCode = ReadlineTextFD.Substring(130, 12),
                                    WFCount = ReadlineTextFD.Substring(143, 2),
                                    ChipCount = ReadlineTextFD.Substring(145, 7),
                                };

                                string ConnectionString = ConfigurationManager.AppSettings["ConnetionStringDBRISTLSI"]!;
                                using (OleDbConnection connection = new OleDbConnection(ConnectionString))
                                {
                                    connection.Open();
                                    using (OleDbCommand sqlCommandCHIPZAIKO = new OleDbCommand("SELECT WFLOTNO FROM CHIPZAIKO WHERE WFLOTNO = ? order by TIMESTAMP desc", connection))
                                    {
                                        sqlCommandCHIPZAIKO.CommandType = CommandType.Text;
                                        sqlCommandCHIPZAIKO.Parameters.AddWithValue("?", waferData.WFLotNo);
                                        using (OleDbDataReader reader = sqlCommandCHIPZAIKO.ExecuteReader())
                                        {
                                            if (reader.HasRows)
                                            {
                                                checkChipzaiko = true;
                                            }
                                            else
                                            {
                                                checkChipzaiko = false;
                                            }
                                        }
                                    }
                                    using (OleDbCommand sqlCommandCHIPNYUKO = new OleDbCommand("SELECT WFLOTNO FROM CHIPNYUKO WHERE WFLOTNO = ? order by TIMESTAMP desc", connection))
                                    {
                                        sqlCommandCHIPNYUKO.CommandType = CommandType.Text;
                                        sqlCommandCHIPNYUKO.Parameters.AddWithValue("?", waferData.WFLotNo);
                                        using (OleDbDataReader reader = sqlCommandCHIPNYUKO.ExecuteReader())
                                        {
                                            if (reader.HasRows)
                                            {
                                                checkChipnyoko = true;
                                            }
                                            else
                                            {
                                                checkChipnyoko = false;
                                            }
                                        }
                                    }
                                }

                                if (checkChipzaiko || checkChipnyoko)
                                {
                                    MainWindow.AlarmBox("This LOT has been uploaded, Please check !!!");
                                    Close();
                                    checkNoInvoiceRows = true;
                                    return;
                                }
                                else
                                {
                                    SetSeq("");
                                    SetWafer(ReadlineTextFD);
                                    STOCKDATA(waferData);
                                    STOCKINDATA(waferData);
                                    UnZipLot(waferData);
                                }
                            }
                        }
                    }
                }
                else
                {
                    MainWindow.AlarmBox("Data not exist check Refidc02.fd !!!");
                }
            }
            catch (Exception ex)
            {
                MainWindow.AlarmBox(ex.Message);
            }
        }
        public static void SetWafer(string getReadlineTextFD)
        {
            int pos = 152;
            int pos2 = 155;
            StringBuilder stringBuilder = new StringBuilder();

            for (int i = 0; i <= 39; i++)
            {
                if (getReadlineTextFD.Substring(pos + (9 * i), 3).Length != 0)
                {
                    string wfSeq = getReadlineTextFD.Substring(pos + (9 * i), 3);
                    string waferChipCount = getReadlineTextFD.Substring(pos2 + (9 * i), 6);

                    // SET_WF_SEQ
                    int wfSeqLength = getReadlineTextFD.Substring(pos + (9 * i), 3).Length;
                    stringBuilder.Append(TmpData);
                    stringBuilder.Append(wfSeqLength switch
                    {
                        1 => "  ",
                        2 => " ",
                        _ => ""
                    });
                    stringBuilder.Append(wfSeq);

                    // SET_CHIP_COUNT
                    int waferChipCountLength = waferChipCount.Trim().Length;
                    stringBuilder.Append(waferChipCountLength switch
                    {
                        1 => "     ",
                        2 => "    ",
                        3 => "   ",
                        4 => "  ",
                        5 => " ",
                        _ => ""
                    });
                    stringBuilder.Append(waferChipCount.Trim());
                }
                else
                {
                    MainWindow.AlarmBox("Can not read data check Refidc02.fd !!!");
                    break;
                }
            }
            resultTmpData = stringBuilder.ToString();
        }
        public static string SetSeq(string setseq)
        {
            try
            {
                int seqNo = 0;
                string? allocatedDate = null;
                string connectionString = ConfigurationManager.AppSettings["ConnetionStringDBRISTLSI"]!;
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                    using (OleDbCommand sqlCommand = new OleDbCommand("SELECT * FROM CHIPSYS WHERE SYSKEY = '01'", connection))
                    {
                        using (OleDbDataReader reader = sqlCommand.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                while (reader.Read())
                                {
                                    seqNo = (int)reader.GetDecimal(2);
                                    allocatedDate = reader.GetString(1);

                                    if (allocatedDate != DateTime.Now.ToString("yyMMdd"))
                                    {
                                        allocatedDate = DateTime.Now.ToString("yyMMdd");
                                        using (OleDbCommand sqlCommandUpdate = new OleDbCommand("UPDATE CHIPSYS " +
                                                                                             "SET ALOCATEDATE = ?, SEQNO = 2 " +
                                                                                             "WHERE SYSKEY = '01'", connection))
                                        {
                                            sqlCommandUpdate.CommandType = CommandType.Text;
                                            sqlCommandUpdate.Parameters.AddWithValue("?", allocatedDate);
                                            sqlCommandUpdate.ExecuteNonQuery();
                                            useqno = "0001";
                                        }
                                    }
                                    else
                                    {
                                        int countSeq = seqNo.ToString().Length;
                                        useqno = countSeq switch
                                        {
                                            4 => seqNo.ToString(),
                                            3 => "0" + seqNo.ToString(),
                                            2 => "00" + seqNo.ToString(),
                                            1 => "000" + seqNo.ToString(),
                                            _ => useqno
                                        };

                                        seqNo++;

                                        using (OleDbCommand sqlCommandUpdate = new OleDbCommand("UPDATE CHIPSYS " +
                                                                     "SET SEQNO = ? " +
                                                                     "WHERE SYSKEY = '01'", connection))
                                        {
                                            sqlCommandUpdate.CommandType = CommandType.Text;
                                            sqlCommandUpdate.Parameters.AddWithValue("?", seqNo);
                                            sqlCommandUpdate.ExecuteNonQuery();
                                        }
                                    }
                                }
                            }
                            else
                            {
                                string sqlCommandInsert = "INSERT INTO CHIPSYS VALUES(?, ?, ?)";
                                try
                                {
                                    allocatedDate = DateTime.Now.ToString("yyMMdd");
                                    using (OleDbCommand sqlCommandQuery = new OleDbCommand(sqlCommandInsert, connection))
                                    {
                                        sqlCommandQuery.CommandType = CommandType.Text;
                                        sqlCommandQuery.Parameters.AddWithValue("?", "01");
                                        sqlCommandQuery.Parameters.AddWithValue("?", allocatedDate);
                                        sqlCommandQuery.Parameters.AddWithValue("?", 2);
                                        sqlCommandQuery.ExecuteNonQuery();
                                        useqno = "0001";
                                    }
                                }
                                catch (OleDbException)
                                {
                                    MainWindow.AlarmBox("Can not connect to the database !!!");
                                }
                            }
                            finseqno = $"Q{allocatedDate!.Substring(1, 5)}{useqno}";
                        }
                    }
                }
            }
            catch (Exception e)
            {
                MainWindow.AlarmBox(e.Message);
            }
            return finseqno!;
        }

        public static void STOCKINDATA(WaferData GetwaferData)
        {
            string TIMESTAMP = DateTime.Now.ToString();
            string STOCKDATE = DateTime.Now.ToString("yyMMdd");
            string WFDATA1 = resultTmpData!.Substring(0, 180);
            string WFDATA2 = resultTmpData!.Substring(180, 180);
            string ReWFDATA1 = WFDATA1.Replace("  0", "   ");
            string ReWFDATA2 = WFDATA2.Replace("  0", "   ");
            string ConnetionString = ConfigurationManager.AppSettings["ConnetionStringDBRISTLSI"]!;
            string sqlInsert = "INSERT INTO CHIPNYUKO (CHIPMODELNAME, MODELCODE1, MODELCODE2, WFLOTNO, SEQNO, OUTDIV," +
                               " RECDIV, STOCKDATE, WFCOUNT, CHIPCOUNT, SLIPNO, SLIPNOEDA, ORDERNO, INVOICENO, HOLDFLAG, CASENO, DELETEFLAG, WFDATA1, WFDATA2, WFINPUT, " +
                               "TIMESTAMP, RFSEQNO) " +
                               "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";
            string sqlUpdate = "UPDATE CHIPNYUKO SET PLASMA = (SELECT PLASMA FROM CHIPMASTER Where CHIPMODELNAME = ? AND PLASMA = '1') " +
                               "WHERE WFLOTNO = ? AND SEQNO = ?";
            using (OleDbConnection connection = new OleDbConnection(ConnetionString))
            {
                try
                {
                    connection.Open();
                    OleDbCommand sqlCommandQuery = new OleDbCommand(sqlInsert, connection);
                    sqlCommandQuery.CommandType = CommandType.Text;
                    sqlCommandQuery.Parameters.AddWithValue("?", GetwaferData.ChipModelName);
                    sqlCommandQuery.Parameters.AddWithValue("?", GetwaferData.ModelCode1);
                    sqlCommandQuery.Parameters.AddWithValue("?", GetwaferData.ModelCode2);
                    sqlCommandQuery.Parameters.AddWithValue("?", GetwaferData.WFLotNo);
                    sqlCommandQuery.Parameters.AddWithValue("?", finseqno);
                    sqlCommandQuery.Parameters.AddWithValue("?", GetwaferData.OutDiv);
                    sqlCommandQuery.Parameters.AddWithValue("?", GetwaferData.RecDiv);
                    sqlCommandQuery.Parameters.AddWithValue("?", STOCKDATE);
                    sqlCommandQuery.Parameters.AddWithValue("?", GetwaferData.WFCount);
                    sqlCommandQuery.Parameters.AddWithValue("?", GetwaferData.ChipCount);
                    sqlCommandQuery.Parameters.AddWithValue("?", "          ");
                    sqlCommandQuery.Parameters.AddWithValue("?", "  ");
                    sqlCommandQuery.Parameters.AddWithValue("?", GetwaferData.OrderNo);
                    sqlCommandQuery.Parameters.AddWithValue("?", GetwaferData.InvoiceNo);
                    sqlCommandQuery.Parameters.AddWithValue("?", "");
                    sqlCommandQuery.Parameters.AddWithValue("?", GetwaferData.CaseNo);
                    sqlCommandQuery.Parameters.AddWithValue("?", "");
                    sqlCommandQuery.Parameters.AddWithValue("?", ReWFDATA1);
                    sqlCommandQuery.Parameters.AddWithValue("?", ReWFDATA2);
                    sqlCommandQuery.Parameters.AddWithValue("?", "1");
                    sqlCommandQuery.Parameters.AddWithValue("?", TIMESTAMP);
                    sqlCommandQuery.Parameters.AddWithValue("?", GetwaferData.RFSeqNo);
                    sqlCommandQuery.ExecuteNonQuery();
                }
                catch (OleDbException)
                {
                    MainWindow.AlarmBox("Can not insert data or connect to the database !!!");
                }
                try
                {
                    OleDbCommand sqlCommandQueryUpdate = new OleDbCommand(sqlUpdate, connection);
                    sqlCommandQueryUpdate.CommandType = CommandType.Text;
                    sqlCommandQueryUpdate.Parameters.AddWithValue("?", GetwaferData.ChipModelName);
                    sqlCommandQueryUpdate.Parameters.AddWithValue("?", GetwaferData.WFLotNo);
                    sqlCommandQueryUpdate.Parameters.AddWithValue("?", finseqno);
                    sqlCommandQueryUpdate.ExecuteNonQuery();
                }
                catch (OleDbException)
                {
                    MainWindow.AlarmBox("Can not update data or connect to the database !!!");
                }
            }
        }
        public static void STOCKDATA(WaferData GetwaferData)
        {
            try
            {
                string STOCKDATE = DateTime.Now.ToString("yyMMdd");
                string TIMESTAMP = DateTime.Now.ToString();
                string ConnectionString = ConfigurationManager.AppSettings["ConnetionStringDBRISTLSI"]!;
                string sqlInsert = "INSERT INTO CHIPZAIKO (CHIPMODELNAME, MODELCODE1, MODELCODE2, WFLOTNO, SEQNO, ENO, LOCATION, WFCOUNT, CHIPCOUNT, STOCKDATE, " +
                                   "RETURNFLAG, REMAINFLAG, HOLDFLAG, STAFFNO, PREOUTFLAG, INVOICENO, PROCESSCODE, DELETEFLAG, TIMESTAMP)" +
                                   "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";
                using (OleDbConnection connection = new OleDbConnection(ConnectionString))
                {
                    connection.Open();
                    OleDbCommand sqlCommandQuery = new OleDbCommand(sqlInsert, connection);
                    sqlCommandQuery.CommandType = CommandType.Text;
                    sqlCommandQuery.Parameters.AddWithValue("?", GetwaferData.ChipModelName);
                    sqlCommandQuery.Parameters.AddWithValue("?", GetwaferData.ModelCode1);
                    sqlCommandQuery.Parameters.AddWithValue("?", GetwaferData.ModelCode2);
                    sqlCommandQuery.Parameters.AddWithValue("?", GetwaferData.WFLotNo);
                    sqlCommandQuery.Parameters.AddWithValue("?", finseqno);
                    sqlCommandQuery.Parameters.AddWithValue("?", "");
                    sqlCommandQuery.Parameters.AddWithValue("?", "");
                    sqlCommandQuery.Parameters.AddWithValue("?", GetwaferData.WFCount);
                    sqlCommandQuery.Parameters.AddWithValue("?", GetwaferData.ChipCount);
                    sqlCommandQuery.Parameters.AddWithValue("?", STOCKDATE);
                    sqlCommandQuery.Parameters.AddWithValue("?", "");
                    sqlCommandQuery.Parameters.AddWithValue("?", "");
                    sqlCommandQuery.Parameters.AddWithValue("?", "");
                    sqlCommandQuery.Parameters.AddWithValue("?", "00001");
                    sqlCommandQuery.Parameters.AddWithValue("?", "");
                    sqlCommandQuery.Parameters.AddWithValue("?", GetwaferData.InvoiceNo);
                    sqlCommandQuery.Parameters.AddWithValue("?", GetwaferData.RecDiv);
                    sqlCommandQuery.Parameters.AddWithValue("?", "");
                    sqlCommandQuery.Parameters.AddWithValue("?", TIMESTAMP);
                    sqlCommandQuery.ExecuteNonQuery();
                }
            }
            catch (OleDbException)
            {
                MainWindow.AlarmBox("Can not insert data or connect to the database !!!");
            }
        }
        public static void UnZipLot(WaferData waferData)
        {
            string extractPath = ConfigurationManager.AppSettings["ExtractPath"]!;
            string checkLotName = ConfigurationManager.AppSettings["ChecklotName"]!; /*Shared Folder*/

            DirectoryInfo directoryInfo = new DirectoryInfo(checkLotName!);
            FileInfo[] files = directoryInfo.GetFiles();

            foreach (FileInfo file in files)
            {
                try
                {
                    string[] splitName = file.Name.Split('.');
                    string lotName = splitName[0];
                    string trimmedLot = waferData.WFLotNo.Replace(" ", "");
                    string lotPathFolder = Path.Combine(extractPath!, trimmedLot);

                    if (lotName == trimmedLot && !file.FullName.EndsWith(".bak"))
                    {
                        if (Directory.Exists(lotPathFolder))
                        {
                            Directory.Delete(lotPathFolder, true);
                        }

                        ZipFile.ExtractToDirectory(file.FullName, lotPathFolder);
                        File.Move(file.FullName, file.FullName + ".bak");
                    }
                }
                catch
                {
                    MainWindow.AlarmBox("Not found zip file in CBAll !!!");
                }
            }
        }
        private void UploadButton_Click(object sender, RoutedEventArgs e)
        {
            bool checkChipzaiko = false;
            bool checkChipnyoko = false;
            if (_InvoiceNo == "        ")
            {
                MainWindow.AlarmConditionBox("Confirm Upload ?");
                if (ModalCondition.setIsyes)
                {
                    UploadDataFDSheet();
                    if (checkNoInvoiceRows)
                    {
                        return;
                    }
                    else
                    {
                        MainWindow.AlarmBox("Upload Successfully");
                        Close();
                    }
                }
            }
            else
            {
                string ConnectionString = ConfigurationManager.AppSettings["ConnetionStringDBRISTLSI"]!;
                using (OleDbConnection connection = new OleDbConnection(ConnectionString))
                {
                    connection.Open();
                    using (OleDbCommand sqlCommandCHIPZAIKO = new OleDbCommand("SELECT INVOICENO FROM CHIPZAIKO where INVOICENO = ? order by TIMESTAMP desc", connection))
                    {
                        sqlCommandCHIPZAIKO.CommandType = CommandType.Text;
                        sqlCommandCHIPZAIKO.Parameters.AddWithValue("?", _InvoiceNo);
                        using (OleDbDataReader reader = sqlCommandCHIPZAIKO.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                checkChipzaiko = true;
                            }
                            else
                            {
                                checkChipzaiko = false;
                            }
                        }
                    }
                    using (OleDbCommand sqlCommandCHIPNYUKO = new OleDbCommand("SELECT INVOICENO FROM CHIPNYUKO WHERE INVOICENO = ? order by TIMESTAMP desc", connection))
                    {
                        sqlCommandCHIPNYUKO.CommandType = CommandType.Text;
                        sqlCommandCHIPNYUKO.Parameters.AddWithValue("?", _InvoiceNo);
                        using (OleDbDataReader reader = sqlCommandCHIPNYUKO.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                checkChipnyoko = true;
                            }
                            else
                            {
                                checkChipnyoko = false;
                            }
                        }
                    }
                }

                if (checkChipzaiko || checkChipnyoko)
                {
                    MainWindow.AlarmBox("This invoice has been uploaded, Please check !!!");
                    Close();
                }
                else
                {
                    UploadDataFDSheet();
                    MainWindow.AlarmBox("Upload Successfully");
                    Close();
                }
            }
        }
    }
}