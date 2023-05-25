using System;
using System.Configuration;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Windows;
using System.Windows.Input;
using static ChipbankImport.ModalFD;

namespace ChipbankImport
{
    public partial class ModalEDSSlip : Window
    {
        public string? zipfileName { get; set; } //from submitButton_Click MainWindow
        public bool checkwaferFail { get; set; } //from ModalDataWafer
        private string? fileName;
        private static string? tmpData;
        private static string? tmpwfLotno;
        private static string? tmp_invoiceNo;
        private static string? tmp_ChipmodelName;
        private static string? waferChipcount;
        private static string? getlotStatus;
        private static string? getplasmaStatus;
        private static string? finseqno;
        private static int sumwfCount;
        private static int sumchipCount;
        private static DataTable dataWFSEQ = new DataTable();
        private static bool isButtoncheckClicked { get; set; }
        private static string? resultTmpData;
        private static StringBuilder stringBuilder = new StringBuilder();
        public ModalEDSSlip()
        {
            InitializeComponent();
            UploadButton.IsEnabled = false;
        }

        private void ExitModalFD_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Card_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                DragMove();
            }
        }
        private void waferLot_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                checkButton_Click(sender, e);
            }
        }
        private void UploadButton_Click(object sender, RoutedEventArgs e)
        {
            if (isButtoncheckClicked)
            {
                UpdateChipnyuko();
                UpdateChipzaiko();
                MoveFile(fileName!);
                MainWindow.AlarmBox("Uploaded Successfully");
                Close();
            }
            else
            {
                MainWindow.AlarmBox("Please click the check button first !!!");
            }
        }
        private void checkButton_Click(object sender, RoutedEventArgs e)
        {
            string waferText = waferLot.Text.ToString().Trim('.');
            if (waferText != "" && waferText.Length == 11 && !waferLot.Text.Contains('.'))
            {
                try
                {
                    string ConnectionString = ConfigurationManager.AppSettings["ConnetionStringDBRISTLSI"]!;
                    string sqlDeleteWF_EDS = "DELETE FROM WF_EDS";
                    string sqlDeleteTMP_EDS = "DELETE FROM TMP_EDS";
                    using (SqlConnection connection = new SqlConnection(ConnectionString))
                    {
                        connection.Open();
                        SqlCommand sqlCommandWF_EDS = new SqlCommand(sqlDeleteWF_EDS, connection);
                        sqlCommandWF_EDS.ExecuteNonQuery();
                        SqlCommand sqlCommandTMP_EDS = new SqlCommand(sqlDeleteTMP_EDS, connection);
                        sqlCommandTMP_EDS.ExecuteNonQuery();
                    }
                }
                catch (SqlException)
                {
                    MainWindow.AlarmBox("Can not connect to the database !!!");
                }
                fileName = $"{waferText}.{zipfileName}.zip";
                Unzip(fileName);
                ShowValues();
            }
            else
            {
                MainWindow.AlarmBox("Barcode Mismatch !!!");
                UploadButton.IsEnabled = false;
            }
        }
        public static void Unzip(string FileName)
        {
            int countLine = 1;
            int line = 0;
            int wfcountFail = 0;
            int WFSEQ;
            sumchipCount = 0;
            sumwfCount = 0;
            string? tmpData1 = null;
            string? tmpData2 = null;
            getplasmaStatus = null;
            tmpData = null;
            string? TMP_ORDERNO = null;
            finseqno = null;
            resultTmpData = null;
            string? FileToCopy = ConfigurationManager.AppSettings["CBOutputPath"] + FileName; /*Shared Folder*/
            string? NewCopyCB = ConfigurationManager.AppSettings["NewCopyCBPath"]!; /*for backup file before sending*/
            string? ProcessPath = ConfigurationManager.AppSettings["ProcessPath"]!;
            string? ChangeDirectory = ConfigurationManager.AppSettings["ChangeDirectory"]!;
            string ConnectionString = ConfigurationManager.AppSettings["ConnetionStringDBRISTLSI"]!;

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
                    MainWindow.AlarmBox("Please check the CBOutput location or the unzip file location !!!");
                }
            }

            Directory.SetCurrentDirectory(ChangeDirectory);

            try
            {
                ZipFile.ExtractToDirectory(NewCopyCB, ProcessPath);
            }
            catch
            {
                MainWindow.AlarmBox("Please check the unzip file location !!!");
            }

            string roblnCsvPath = Path.Combine(ProcessPath, "ROBIN_L.CSV");
            if (File.Exists(roblnCsvPath))
            {
                string[] csvLines = File.ReadAllLines(roblnCsvPath);
                foreach (string lineText in csvLines)
                {
                    if (countLine == 1)
                    {
                        tmp_ChipmodelName = lineText;
                    }
                    if (countLine == 7)
                    {
                        tmp_invoiceNo = lineText.Trim();
                    }
                    countLine++;
                }
            }
            else
            {
                MainWindow.AlarmBox("Data not exist check ROBIN_L.CSV !!!");
            }

            string cbInputPath = Path.Combine(ProcessPath, "CBinput.txt");
            if (File.Exists(cbInputPath))
            {
                string[] cbInputLines = File.ReadAllLines(cbInputPath);
                foreach (string lineText in cbInputLines)
                {
                    line++;
                    if (line == 2)
                    {
                        tmpwfLotno = lineText;
                    }
                    if (line >= 3)
                    {
                        WFSEQ = line - 2;
                        waferChipcount = lineText;
                        if (waferChipcount != "")
                        {
                            if (waferChipcount == "0")
                            {
                                wfcountFail++;
                            }
                            else
                            {
                                if (WFSEQ.ToString().Length == 1)
                                {
                                    resultTmpData = stringBuilder.Append(tmpData).Append("  ").Append(WFSEQ).ToString();
                                }
                                else if (WFSEQ.ToString().Length == 2)
                                {
                                    resultTmpData = stringBuilder.Append(tmpData).Append(' ').Append(WFSEQ).ToString();
                                }
                                else if (WFSEQ.ToString().Length == 3)
                                {
                                    resultTmpData = stringBuilder.Append(tmpData).Append(WFSEQ).ToString();
                                }
                                if (!string.IsNullOrEmpty(waferChipcount))
                                {
                                    sumchipCount += int.Parse(waferChipcount);
                                    if (waferChipcount!.Length == 1)
                                    {
                                        resultTmpData = stringBuilder.Append(tmpData).Append("     ").Append(waferChipcount).ToString();
                                    }
                                    else if (waferChipcount!.Length == 2)
                                    {
                                        resultTmpData = stringBuilder.Append(tmpData).Append("    ").Append(waferChipcount).ToString();
                                    }
                                    else if (waferChipcount!.Length == 3)
                                    {
                                        resultTmpData = stringBuilder.Append(tmpData).Append("   ").Append(waferChipcount).ToString();
                                    }
                                    else if (waferChipcount!.Length == 4)
                                    {
                                        resultTmpData = stringBuilder.Append(tmpData).Append("  ").Append(waferChipcount).ToString();
                                    }
                                    else if (waferChipcount!.Length == 5)
                                    {
                                        resultTmpData = stringBuilder.Append(tmpData).Append(' ').Append(waferChipcount).ToString();
                                    }
                                    else if (waferChipcount!.Length == 6)
                                    {
                                        resultTmpData = stringBuilder.Append(tmpData).Append(waferChipcount).ToString();
                                    }
                                    if (waferChipcount != "0")
                                    {
                                        sumwfCount++;
                                        string sqlInsert = "INSERT INTO WF_EDS (WFLOTNO,WFSEQ,CHIPCOUNT,INVOICENO) " +
                                                           "VALUES (@WFLOTNO,@WFSEQ,@CHIPCOUNT,@INVOICENO)";
                                        using (SqlConnection connection = new SqlConnection(ConnectionString))
                                        {
                                            connection.Open();
                                            SqlCommand sqlCommandQuery = new SqlCommand(sqlInsert, connection);
                                            sqlCommandQuery.Parameters.AddWithValue("@WFLOTNO", tmpwfLotno);
                                            sqlCommandQuery.Parameters.AddWithValue("@WFSEQ", WFSEQ);
                                            sqlCommandQuery.Parameters.AddWithValue("@CHIPCOUNT", waferChipcount);
                                            sqlCommandQuery.Parameters.AddWithValue("@INVOICENO", tmp_invoiceNo);
                                            sqlCommandQuery.ExecuteNonQuery();
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                if (resultTmpData!.Length <= 180)
                {
                    tmpData1 = resultTmpData!.Substring(0, Math.Min(resultTmpData.Length, 180));
                    tmpData2 = "";
                }
                else
                {
                    tmpData1 = resultTmpData!.Substring(0, Math.Min(resultTmpData.Length, 180));
                    tmpData2 = resultTmpData!.Substring(180, Math.Max(0, Math.Min(resultTmpData.Length - 180, 180)));
                }
            }

            ConnectionString = ConfigurationManager.AppSettings["ConnetionStringDBRISTLSI"]!;
            string ConnetionStringMapOnline = ConfigurationManager.AppSettings["ConnetionStringMapOnline"]!;
            string sqlSelectWFSEQ = "SELECT WFSEQ, CHIPCOUNT FROM WF_EDS ORDER BY WFSEQ";
            string sqlSelectCHIPMASTER = "SELECT * FROM CHIPMASTER WHERE CHIPMODELNAME = @tmp_ChipmodelName";
            string sqlselectSample = "SELECT DeviceName, LotNo, FlagLastShipout FROM MapOnline.dbo.EDSFlow " +
                                     "WHERE DeviceName LIKE '%8' AND LotNo = @LotNo AND FlowName = 'output' AND FlagLastShipout = 1";

            using (SqlConnection connection = new SqlConnection(ConnectionString))
            {
                connection.Open();
                SqlCommand sqlCommandQueryChip = new SqlCommand(sqlSelectCHIPMASTER, connection);
                sqlCommandQueryChip.Parameters.AddWithValue("@tmp_ChipmodelName", tmp_ChipmodelName);

                using (SqlDataReader readerQueryChip = sqlCommandQueryChip.ExecuteReader())
                {
                    if (readerQueryChip.HasRows)
                    {
                        foreach (DbDataRecord record in readerQueryChip)
                        {
                            getplasmaStatus = (string?)record.GetValue(1);
                        }
                    }
                    else
                    {
                        getplasmaStatus = " ";
                    }
                }

                SqlCommand sqlCommandQueryWF = new SqlCommand(sqlSelectWFSEQ, connection);

                using (SqlDataReader readerQueryWF = sqlCommandQueryWF.ExecuteReader())
                {
                    if (readerQueryWF.HasRows)
                    {
                        dataWFSEQ = new DataTable();
                        for (int i = 0; i < readerQueryWF.FieldCount; i++)
                        {
                            dataWFSEQ.Columns.Add(readerQueryWF.GetName(i), readerQueryWF.GetFieldType(i));
                        }

                        while (readerQueryWF.Read())
                        {
                            DataRow row = dataWFSEQ.NewRow();

                            for (int i = 0; i < readerQueryWF.FieldCount; i++)
                            {
                                row[i] = readerQueryWF[i];
                            }
                            dataWFSEQ.Rows.Add(row);
                        }
                    }
                }
            }

            using (SqlConnection connectionMapOnline = new SqlConnection(ConnetionStringMapOnline))
            {
                connectionMapOnline.Open();
                SqlCommand sqlCommandQuerySample = new SqlCommand(sqlselectSample, connectionMapOnline);
                sqlCommandQuerySample.Parameters.AddWithValue("@LotNo", tmpwfLotno);

                using (SqlDataReader readerQuerysample = sqlCommandQuerySample.ExecuteReader())
                {
                    if (readerQuerysample.HasRows)
                    {
                        getlotStatus = "Sample";
                    }
                    else
                    {
                        getlotStatus = "Mass Production";
                    }
                }
            }

            string TMP_CASENO = "      ";
            string TMP_OUTDIV = "QI830";
            string TMP_RECDIV = "TI970";
            string checkDigit = tmpwfLotno!.Substring(0, 1);

            if (Char.IsDigit(checkDigit[0]))
            {
                TMP_ORDERNO = "E" + tmpwfLotno;
            }
            else
            {
                TMP_ORDERNO = tmpwfLotno;
            }

            SetSeq();

            string sqlInsertTMP_EDS = "INSERT INTO TMP_EDS (CHIPMODELNAME,WFLOTNO,WFCOUNT,CHIPCOUNT,INVOICENO,CASENO,OUTDIV,RECDIV,ORDERNO,PLASMA,WFDATA1,WFDATA2,SEQNO,WFCOUNT_FAIL) " +
                                      "VALUES (@CHIPMODELNAME, @WFLOTNO, @WFCOUNT, @CHIPCOUNT, @INVOICENO, @CASENO, @OUTDIV, @RECDIV, @ORDERNO, @PLASMA, @WFDATA1, @WFDATA2, @SEQNO, @WFCOUNT_FAIL)";

            string ALOCATEDATE = DateTime.Now.ToString("yyMMdd");
            string useqno = "0001";
            finseqno = $"Q{ALOCATEDATE!.Substring(1, 5)}{useqno}";

            using (SqlConnection connection = new SqlConnection(ConnectionString))
            {
                connection.Open();
                SqlCommand sqlCommandQuery = new SqlCommand(sqlInsertTMP_EDS, connection);
                sqlCommandQuery.Parameters.AddWithValue("@CHIPMODELNAME", tmp_ChipmodelName);
                sqlCommandQuery.Parameters.AddWithValue("@WFLOTNO", tmpwfLotno);
                sqlCommandQuery.Parameters.AddWithValue("@WFCOUNT", sumwfCount);
                sqlCommandQuery.Parameters.AddWithValue("@CHIPCOUNT", sumchipCount);
                sqlCommandQuery.Parameters.AddWithValue("@INVOICENO", tmp_invoiceNo);
                sqlCommandQuery.Parameters.AddWithValue("@CASENO", TMP_CASENO);
                sqlCommandQuery.Parameters.AddWithValue("@OUTDIV", TMP_OUTDIV);
                sqlCommandQuery.Parameters.AddWithValue("@RECDIV", TMP_RECDIV);
                sqlCommandQuery.Parameters.AddWithValue("@ORDERNO", TMP_ORDERNO);
                sqlCommandQuery.Parameters.AddWithValue("@PLASMA", getplasmaStatus);
                sqlCommandQuery.Parameters.AddWithValue("@WFDATA1", tmpData1);
                sqlCommandQuery.Parameters.AddWithValue("@WFDATA2", tmpData2);
                sqlCommandQuery.Parameters.AddWithValue("@SEQNO", finseqno);
                sqlCommandQuery.Parameters.AddWithValue("@WFCOUNT_FAIL", wfcountFail);
                sqlCommandQuery.ExecuteNonQuery();
            }
        }

        private void ShowValues()
        {
            if (tmp_invoiceNo != null && finseqno != null)
            {
                SetInvoiceValues();
                SetPlasmaStatus();
                ShowDataWafer();
                if (checkwaferFail)
                {
                    UploadButton.IsEnabled = false;
                }
                isButtoncheckClicked = true;
                UploadButton.IsEnabled = true;
            }
            else
            {
                ClearInvoiceValues();
                UploadButton.IsEnabled = false;
            }
        }
        private void SetInvoiceValues()
        {
            invoiceNo.Text = tmp_invoiceNo;
            chipModel.Text = tmp_ChipmodelName;
            waferCount.Text = sumwfCount.ToString();
            chipCount.Text = sumchipCount.ToString();
            lotStatus.Text = getlotStatus;
            seqNo.Text = finseqno;
        }
        private void SetPlasmaStatus()
        {
            if (getplasmaStatus == "1")
            {
                plasmaStatus.Text = "Plasma";
            }
            else if (getplasmaStatus == "0")
            {
                plasmaStatus.Text = "No Plasma";
            }
            else
            {
                plasmaStatus.Text = "No Data";
            }
        }
        private void ShowDataWafer()
        {
            DataWafer dataWafer = new DataWafer();
            dataWafer.tmpwfLotno = tmpwfLotno;
            dataWafer.sumwfCount = sumwfCount;
            dataWafer.sumchipCount = sumchipCount;
            dataWafer.dataGridwafer.ItemsSource = dataWFSEQ.DefaultView;
            dataWafer.ShowDialog();
            checkwaferFail = dataWafer.checkwaferFail;
        }
        private void ClearInvoiceValues()
        {
            invoiceNo.Text = null;
            chipModel.Text = null;
            waferCount.Text = null;
            chipCount.Text = null;
            lotStatus.Text = null;
            seqNo.Text = null;
            plasmaStatus.Text = null;
        }
        private static void UpdateChipnyuko()
        {
            string ConnectionString = ConfigurationManager.AppSettings["ConnetionStringDBRISTLSI"]!;
            string sqlSelectTMPEDS = "SELECT* FROM TMP_EDS";
            using (SqlConnection connection = new SqlConnection(ConnectionString))
            {
                connection.Open();
                using (SqlCommand sqlCommandQuery = new SqlCommand(sqlSelectTMPEDS, connection))
                {
                    SqlDataReader reader = sqlCommandQuery.ExecuteReader();
                    while (reader.Read())
                    {
                        string CHIPMODELNAME = reader.GetString(0);
                        string WFLOTNO = reader.GetString(1);
                        int WFCOUNT = (int)reader.GetSqlDecimal(2);
                        int CHIPCOUNT = (int)reader.GetSqlDecimal(3);
                        string INVOICENO = reader.GetString(4);
                        string CASENO = reader.GetString(5);
                        string OUTDIV = reader.GetString(6);
                        string RECDIV = reader.GetString(7);
                        string ORDERNO = reader.GetString(8);
                        string WFDATA1 = reader.GetString(9);
                        string WFDATA2 = reader.GetString(10);
                        string SEQNO = reader.GetString(11);
                        string PLASMA = reader.GetString(12);
                        int WFCOUNT_FAIL = (int)reader.GetSqlDecimal(13);

                        string insertQuery = "INSERT INTO CHIPNYUKO (CHIPMODELNAME, MODELCODE1, MODELCODE2, WFLOTNO, SEQNO, ENO, RFSEQNO, OUTDIV, RECDIV, STOCKDATE, RETURNFLAG, WFCOUNT,CHIPCOUNT, " +
                                             "ORDERNO, HIGHREL, AGARIDATE, BUNKAN, RETURNCLASS, SLIPNO, SLIPNOEDA,STAFFNO, WFINPUT, TESTERNO, PROGRAMVER,RECYCLE, RINGNO, INVOICENO, HOLDFLAG,CASENO, " +
                                             "DIRECTCLASS, DELETEFLAG, WFDATA1, WFDATA2, TIMESTAMP, MOVE_ORDERNO, PLASMA,WFCOUNT_FAIL) " +
                                             "VALUES (@CHIPMODELNAME, @MODELCODE1, @MODELCODE2, @WFLOTNO, @SEQNO, @ENO, @RFSEQNO, @OUTDIV, @RECDIV, @STOCKDATE, @RETURNFLAG, @WFCOUNT,@CHIPCOUNT, @ORDERNO, @HIGHREL, " +
                                             "@AGARIDATE, @BUNKAN, @RETURNCLASS, @SLIPNO, @SLIPNOEDA, @STAFFNO, @WFINPUT, @TESTERNO, @PROGRAMVER, @RECYCLE, @RINGNO, @INVOICENO, @HOLDFLAG, @CASENO, @DIRECTCLASS, @DELETEFLAG, " +
                                             "@WFDATA1, @WFDATA2, @TIMESTAMP, @MOVE_ORDERNO, @PLASMA, @WFCOUNT_FAIL)";
                        using (SqlCommand insertCommand = new SqlCommand(insertQuery, connection))
                        {
                            insertCommand.Parameters.AddWithValue("@CHIPMODELNAME", CHIPMODELNAME);
                            insertCommand.Parameters.AddWithValue("@MODELCODE1", "");
                            insertCommand.Parameters.AddWithValue("@MODELCODE2", "");
                            insertCommand.Parameters.AddWithValue("@WFLOTNO", WFLOTNO);
                            insertCommand.Parameters.AddWithValue("@SEQNO", SEQNO);
                            insertCommand.Parameters.AddWithValue("@ENO", "");
                            insertCommand.Parameters.AddWithValue("@RFSEQNO", "");
                            insertCommand.Parameters.AddWithValue("@OUTDIV", OUTDIV);
                            insertCommand.Parameters.AddWithValue("@RECDIV", RECDIV);
                            insertCommand.Parameters.AddWithValue("@STOCKDATE", DateTime.Now.ToString("yyMMdd"));
                            insertCommand.Parameters.AddWithValue("@RETURNFLAG", "");
                            insertCommand.Parameters.AddWithValue("@WFCOUNT", WFCOUNT);
                            insertCommand.Parameters.AddWithValue("@CHIPCOUNT", CHIPCOUNT);
                            insertCommand.Parameters.AddWithValue("@ORDERNO", ORDERNO);
                            insertCommand.Parameters.AddWithValue("@HIGHREL", "");
                            insertCommand.Parameters.AddWithValue("@AGARIDATE", DateTime.Now.ToString("yyMMdd"));
                            insertCommand.Parameters.AddWithValue("@BUNKAN", "");
                            insertCommand.Parameters.AddWithValue("@RETURNCLASS", "");
                            insertCommand.Parameters.AddWithValue("@SLIPNO", "");
                            insertCommand.Parameters.AddWithValue("@SLIPNOEDA", "");
                            insertCommand.Parameters.AddWithValue("@STAFFNO", "001");
                            insertCommand.Parameters.AddWithValue("@WFINPUT", "");
                            insertCommand.Parameters.AddWithValue("@TESTERNO", "");
                            insertCommand.Parameters.AddWithValue("@PROGRAMVER", "");
                            insertCommand.Parameters.AddWithValue("@RECYCLE", "");
                            insertCommand.Parameters.AddWithValue("@RINGNO", "");
                            insertCommand.Parameters.AddWithValue("@INVOICENO", INVOICENO);
                            insertCommand.Parameters.AddWithValue("@HOLDFLAG", "");
                            insertCommand.Parameters.AddWithValue("@CASENO", CASENO);
                            insertCommand.Parameters.AddWithValue("@DIRECTCLASS", "");
                            insertCommand.Parameters.AddWithValue("@DELETEFLAG", "");
                            insertCommand.Parameters.AddWithValue("@WFDATA1", WFDATA1);
                            insertCommand.Parameters.AddWithValue("@WFDATA2", WFDATA2);
                            insertCommand.Parameters.AddWithValue("@TIMESTAMP", DateTime.Now.ToString());
                            insertCommand.Parameters.AddWithValue("@MOVE_ORDERNO", "");
                            insertCommand.Parameters.AddWithValue("@PLASMA", PLASMA);
                            insertCommand.Parameters.AddWithValue("@WFCOUNT_FAIL", WFCOUNT_FAIL);
                            insertCommand.ExecuteNonQuery();
                        }
                    }
                    reader.Close();
                }
            }
        }
        private static void UpdateChipzaiko()
        {
            string ConnectionString = ConfigurationManager.AppSettings["ConnetionStringDBRISTLSI"]!;
            string sqlSelectTMPEDS = "SELECT* FROM TMP_EDS";
            using (SqlConnection connection = new SqlConnection(ConnectionString))
            {
                connection.Open();
                using (SqlCommand sqlCommandQuery = new SqlCommand(sqlSelectTMPEDS, connection))
                {
                    SqlDataReader reader = sqlCommandQuery.ExecuteReader();
                    while (reader.Read())
                    {
                        string CHIPMODELNAME = reader.GetString(0);
                        string WFLOTNO = reader.GetString(1);
                        int WFCOUNT = (int)reader.GetSqlDecimal(2);
                        int CHIPCOUNT = (int)reader.GetSqlDecimal(3);
                        string SEQNO = reader.GetString(11);
                        string insertQuery = "INSERT INTO CHIPZAIKO (CHIPMODELNAME, MODELCODE1, MODELCODE2, WFLOTNO, SEQNO, ENO, LOCATION, WFCOUNT, CHIPCOUNT, STOCKDATE, RETURNFLAG, REMAINFLAG, HOLDFLAG, STAFFNO, PREOUTFLAG, INVOICENO, PROCESSCODE, DELETEFLAG, TIMESTAMP) " +
                        "VALUES (@CHIPMODELNAME, @MODELCODE1, @MODELCODE2, @WFLOTNO, @SEQNO, @ENO, @LOCATION, @WFCOUNT, @CHIPCOUNT, @STOCKDATE, @RETURNFLAG, @REMAINFLAG, @HOLDFLAG, @STAFFNO, @PREOUTFLAG, @INVOICENO, @PROCESSCODE, @DELETEFLAG, @TIMESTAMP)";
                        using (SqlCommand insertCommand = new SqlCommand(insertQuery, connection))
                        {
                            insertCommand.Parameters.AddWithValue("@CHIPMODELNAME", CHIPMODELNAME);
                            insertCommand.Parameters.AddWithValue("@MODELCODE1", "");
                            insertCommand.Parameters.AddWithValue("@MODELCODE2", "");
                            insertCommand.Parameters.AddWithValue("@WFLOTNO", WFLOTNO);
                            insertCommand.Parameters.AddWithValue("@SEQNO", SEQNO);
                            insertCommand.Parameters.AddWithValue("@ENO", "");
                            insertCommand.Parameters.AddWithValue("@LOCATION", "");
                            insertCommand.Parameters.AddWithValue("@WFCOUNT", WFCOUNT);
                            insertCommand.Parameters.AddWithValue("@CHIPCOUNT", CHIPCOUNT);
                            insertCommand.Parameters.AddWithValue("@STOCKDATE", DateTime.Now.ToString("yyMMdd"));
                            insertCommand.Parameters.AddWithValue("@RETURNFLAG", "");
                            insertCommand.Parameters.AddWithValue("@REMAINFLAG", "");
                            insertCommand.Parameters.AddWithValue("@HOLDFLAG", "");
                            insertCommand.Parameters.AddWithValue("@STAFFNO", "001");
                            insertCommand.Parameters.AddWithValue("@PREOUTFLAG", "");
                            insertCommand.Parameters.AddWithValue("@INVOICENO", "");
                            insertCommand.Parameters.AddWithValue("@PROCESSCODE", "TI970");
                            insertCommand.Parameters.AddWithValue("@DELETEFLAG", "");
                            insertCommand.Parameters.AddWithValue("@TIMESTAMP", DateTime.Now.ToString());
                            insertCommand.ExecuteNonQuery();
                        }
                    }
                    reader.Close();
                }
            }
        }
        private void MoveFile(string getFileName)
        {
            string? FileToCopy = ConfigurationManager.AppSettings["CBOutputPath"] + getFileName; /*Shared Folder*/
            string? CopytoPath = ConfigurationManager.AppSettings["CopytoPath"]!; /*Backup Folder*/
            string? ExtractPath = ConfigurationManager.AppSettings["ExtractPath"]!; /*Extract Folder*/
            string destinationFilePath = System.IO.Path.Combine(CopytoPath, fileName!);
            string fileamenotExtension = System.IO.Path.GetFileNameWithoutExtension(FileToCopy);
            int fileLotname = fileamenotExtension.IndexOf('.');
            string LotNo = fileamenotExtension.Substring(0, fileLotname);
            if (File.Exists(FileToCopy))
            {
                try
                {
                    ZipFile.ExtractToDirectory(FileToCopy, ExtractPath + LotNo);
                    File.Move(FileToCopy, destinationFilePath, true);
                }
                catch (Exception)
                {
                    MainWindow.AlarmBox("Not found zip file in the CBOutput location !!!");
                }
            }
            else
            {
                MainWindow.AlarmBox("Not found zip file in the CBOutput location !!!");
            }
        }
    }
}