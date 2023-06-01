using System;
using System.Configuration;
using System.Data.SqlClient;
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
                                SetWafer(ReadlineTextFD);
                                STOCKDATA(waferData);
                                STOCKINDATA(waferData);
                                UnZipLot(waferData);
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
        public static void SetSeq()
        {
            try
            {
                int seqNo = 0;
                string? allocatedDate = null;
                string connectionString = ConfigurationManager.AppSettings["ConnetionStringDBRISTLSI"]!;
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    using (SqlCommand sqlCommand = new SqlCommand("SELECT * FROM CHIPSYS WHERE SYSKEY = @SYSKEY", connection))
                    {
                        sqlCommand.Parameters.AddWithValue("@SYSKEY", "01");
                        using (SqlDataReader reader = sqlCommand.ExecuteReader())
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
                                        using (SqlCommand sqlCommandUpdate = new SqlCommand("UPDATE CHIPSYS " +
                                                                                             "SET ALOCATEDATE = @ALOCATEDATE, SEQNO = 2 " +
                                                                                             "WHERE SYSKEY = '01'", connection))
                                        {
                                            sqlCommandUpdate.Parameters.AddWithValue("@ALOCATEDATE", allocatedDate);
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

                                        using (SqlCommand sqlCommandUpdate = new SqlCommand("UPDATE CHIPSYS " +
                                                                     "SET SEQNO = @SEQNO " +
                                                                     "WHERE SYSKEY = '01'", connection))
                                        {
                                            sqlCommandUpdate.Parameters.AddWithValue("@SEQNO", seqNo);
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
                                    allocatedDate = DateTime.Now.ToString("yyMMdd");
                                    using (SqlCommand sqlCommandQuery = new SqlCommand(sqlCommandInsert, connection))
                                    {
                                        sqlCommandQuery.Parameters.AddWithValue("@SYSKEY", "01");
                                        sqlCommandQuery.Parameters.AddWithValue("@ALOCATEDATE", allocatedDate);
                                        sqlCommandQuery.Parameters.AddWithValue("@SEQNO", 2);
                                        sqlCommandQuery.ExecuteNonQuery();
                                        useqno = "0001";
                                    }
                                }
                                catch (SqlException)
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
                               "VALUES (@CHIPMODELNAME, @MODELCODE1, @MODELCODE2, @WFLOTNO, @SEQNO, @OUTDIV, @RECDIV, @STOCKDATE, @WFCOUNT, " +
                               "@CHIPCOUNT, @SLIPNO, @SLIPNOEDA, @ORDERNO, @INVOICENo, @HOLDFLAG, @CASENO, @DELETEFLAG, @tmpdata1, @tmpdata2, @WFINPUT, @TIMESTAMP, @RF_SEQNO)";
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
                    sqlCommandQuery.Parameters.AddWithValue("@DELETEFLAG", "");
                    sqlCommandQuery.Parameters.AddWithValue("@tmpdata1", ReWFDATA1);
                    sqlCommandQuery.Parameters.AddWithValue("@tmpdata2", ReWFDATA2);
                    sqlCommandQuery.Parameters.AddWithValue("@WFINPUT", "1");
                    sqlCommandQuery.Parameters.AddWithValue("@TIMESTAMP", TIMESTAMP);
                    sqlCommandQuery.Parameters.AddWithValue("@RF_SEQNO", GetwaferData.RFSeqNo);
                    sqlCommandQuery.ExecuteNonQuery();
                }
                catch (SqlException)
                {
                    MainWindow.AlarmBox("Can not insert data or connect to the database !!!");
                }
                try
                {
                    SqlCommand sqlCommandQueryUpdate = new SqlCommand(sqlUpdate, connection);
                    sqlCommandQueryUpdate.Parameters.AddWithValue("@CHIPMODELNAME", GetwaferData.ChipModelName);
                    sqlCommandQueryUpdate.Parameters.AddWithValue("@WFLOTNO", GetwaferData.WFLotNo);
                    sqlCommandQueryUpdate.Parameters.AddWithValue("@finseqno", finseqno);
                    sqlCommandQueryUpdate.ExecuteNonQuery();
                }
                catch (SqlException)
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
            try
            {
                if (_InvoiceNo == "        ")
                {
                    MainWindow.AlarmConditionBox("Confirm Upload ?");
                    if (ModalCondition.setIsyes)
                    {
                        SetSeq();
                        UploadDataFDSheet();
                        MainWindow.AlarmBox("Upload Successfully");
                        Close();
                    }
                }
                else
                {
                    string ConnectionString = ConfigurationManager.AppSettings["ConnetionStringDBRISTLSI"]!;
                    using (SqlConnection connection = new SqlConnection(ConnectionString))
                    {
                        connection.Open();
                        using (SqlCommand sqlCommand = new SqlCommand("SELECT INVOICENO FROM CHIPZAIKO where INVOICENO = @_InvoiceNo order by TIMESTAMP desc", connection))
                        {
                            sqlCommand.Parameters.AddWithValue("@_InvoiceNo", _InvoiceNo);
                            using (SqlDataReader reader = sqlCommand.ExecuteReader())
                            {
                                if (reader.HasRows)
                                {
                                    MainWindow.AlarmBox("This invoice has been uploaded, Please check !!!");
                                    Close();
                                }
                                else
                                {
                                    SetSeq();
                                    UploadDataFDSheet();
                                    MainWindow.AlarmBox("Upload Successfully");
                                    Close();
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MainWindow.AlarmBox(ex.Message);
            }
        }
    }
}
