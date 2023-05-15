using ChipbankImport;
using System;
using System.Configuration;
using System.Data.Common;
using System.Data.SqlClient;
using System.IO;
using System.IO.Compression;
using System.Windows;
using System.Windows.Input;

namespace ChipbankImport
{
    public partial class ModalFD : Window
    {
        public string? _InvoiceNo { get; set; } // from Mainwindow
        public int _LotCount { get; set; } // from Mainwindow
        public string? GetProcessPath { get; set; } // from Mainwindow
        private static string? TmpData;
        private static string? useqno;
        private static string? finseqno;
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
        private void UploadDataFDSheet(bool checkError)
        {
            string? ReadlineTextFD = null;
            if (File.Exists(GetProcessPath))
            {
                using (FileStream fileStream = new FileStream(GetProcessPath, FileMode.Open))
                {
                    if (checkError == true)
                    {
                        fileStream.Close();
                        fileStream.Dispose();
                    }
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
                            SET_WAFER(ReadlineTextFD);
                            STOCKDATA(waferData);
                            STOCKINDATA(waferData);
                            UnZipLOT(waferData);
                        }
                    }
                }
            }
            else
            {
                MainWindow.AlarmBox("Data not exist check Refidc02.fd !!!");
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
                    MainWindow.AlarmBox("Can not read data check Refidc02.fd !!!");
                    break;
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
                                            //sqlCommandUpdate.ExecuteNonQuery();
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
                                            // sqlCommandUpdate.ExecuteNonQuery();
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
                                    //   sqlCommandQuery.ExecuteNonQuery();
                                    useqno = "0001";
                                }
                                catch (SqlException)
                                {
                                    MainWindow.AlarmBox("Can not connect to the database !!!");
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
                MainWindow.AlarmBox(e.Message);
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
                    // sqlCommandQuery.ExecuteNonQuery();
                }
                catch (SqlException)
                {
                    MainWindow.AlarmBox("Can not insert data or connect to the database !!!");
                }
                try
                {
                    SqlCommand sqlCommandQuery = new SqlCommand(sqlUpdate, connection);
                    sqlCommandQuery.Parameters.AddWithValue("@CHIPMODELNAME", GetwaferData.ChipModelName);
                    sqlCommandQuery.Parameters.AddWithValue("@WFLOTNO", GetwaferData.WFLotNo);
                    sqlCommandQuery.Parameters.AddWithValue("@finseqno", finseqno);
                    //  sqlCommandQuery.ExecuteNonQuery();
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
                    // sqlCommandQuery.ExecuteNonQuery();
                }
            }
            catch (SqlException)
            {
                MainWindow.AlarmBox("Can not insert data or connect to the database !!!");
            }
        }
        public void UnZipLOT(WaferData GetwaferData)
        {
            string? extractPath = ConfigurationManager.AppSettings["ExtractPath"];
            string? ChecklotName = ConfigurationManager.AppSettings["ChecklotName"]; /*Shared Folder*/
            DirectoryInfo directoryInfo = new DirectoryInfo(ChecklotName!);
            FileInfo[] files = directoryInfo.GetFiles();

            foreach (FileInfo file in files)
            {
                try
                {
                    string[] splitName = file.Name.Split('.');
                    string lotName = splitName[0];
                    string trimLot = GetwaferData.WFLotNo.Replace(" ", "");
                    string LotpathFolder = extractPath! + trimLot;
                    if (lotName == trimLot)
                    {
                        if (Directory.Exists(LotpathFolder))
                        {
                            Directory.Delete(LotpathFolder, true);
                            ZipFile.ExtractToDirectory(file.FullName, LotpathFolder);
                        }
                        else
                        {
                            ZipFile.ExtractToDirectory(file.FullName, LotpathFolder);
                        }
                    }
                }
                catch
                {
                    MainWindow.AlarmBox("Not found zip file in CBAll !!!");
                    UploadDataFDSheet(true);
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
                    if (ModalCondition.setIsyes == true)
                    {
                        SETSEQ();
                        UploadDataFDSheet(false);
                        MainWindow.AlarmBox("Upload Successfully");
                        Close();
                    }
                    else
                    {
                        Close();
                    }
                }
                else
                {
                    SETSEQ();
                    UploadDataFDSheet(false);
                    MainWindow.AlarmBox("Upload Successfully");
                    Close();
                }
            }
            catch (Exception ex)
            {
                MainWindow.AlarmBox(ex.Message);
            }
        }
    }
}
