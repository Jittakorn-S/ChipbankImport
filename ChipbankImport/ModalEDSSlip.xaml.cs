using System;
using System.Configuration;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Input;
using static ChipbankImport.ModalFD;

namespace ChipbankImport
{
    public partial class ModalEDSSlip : Window
    {
        private static DataTable dataWFSEQ = new DataTable();
        private static StringBuilder stringBuilder = new StringBuilder();
        private static bool checkHasrowsNyuko;
        private static bool checkHasrowsZaiko;
        private static bool isButtoncheckClicked { get; set; }
        private static int sumchipCount;
        private static int sumwfCount;
        private static string? checkWFLOTNOzaiko;
        private static string? checkWFLOTNOnyuko;
        private static string? getlotStatus;
        private static string? getplasmaStatus;
        private static string? resultTmpData;
        private static string? tmpData;
        private static string? tmp_ChipmodelName;
        private static string? tmp_invoiceNo;
        private static string? tmpwfLotno;
        private static string? waferChipcount;
        private static string? fileName;
        private static string? getfinseqno;
        public bool checkwaferFail { get; set; } //from ModalDataWafer
        public string? zipfileName { get; set; } //from submitButton_Click MainWindow

        public ModalEDSSlip()
        {
            InitializeComponent();
            UploadButton.IsEnabled = false;
            waferLot.Focus();
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
                if (getlotStatus == "Sample")
                {
                    MoveFile(fileName!);
                    Close();
                }
                else
                {
                    UpdateChipnyuko();
                    UpdateChipzaiko();
                    if (checkHasrowsNyuko || checkHasrowsZaiko)
                    {
                        if (checkWFLOTNOnyuko != null)
                        {
                            UnzipSplitLot(checkWFLOTNOnyuko, fileName!);
                            Close();
                        }
                        else if (checkWFLOTNOzaiko != null)
                        {
                            UnzipSplitLot(checkWFLOTNOzaiko, fileName!);
                            Close();
                        }
                        else
                        {
                            MainWindow.AlarmBox("This lot has no data !!!");
                            Close();
                        }
                    }
                    else
                    {
                        MoveFile(fileName!);
                        Close();
                    }
                }
            }
            else
            {
                MainWindow.AlarmBox("Please click the check button first !!!");
            }
        }
        private void checkButton_Click(object sender, RoutedEventArgs e)
        {
            string waferText = waferLot.Text.ToString().Trim('.');
            if (waferText != "" && waferText.Length <= 12)
            {
                try
                {
                    string ConnectionString = ConfigurationManager.AppSettings["ConnetionStringDBRISTLSI"]!;
                    string sqlDeleteWF_EDS = "DELETE FROM WF_EDS";
                    string sqlDeleteTMP_EDS = "DELETE FROM TMP_EDS";
                    using (OleDbConnection connection = new OleDbConnection(ConnectionString))
                    {
                        connection.Open();
                        OleDbCommand sqlCommandWF_EDS = new OleDbCommand(sqlDeleteWF_EDS, connection);
                        sqlCommandWF_EDS.ExecuteNonQuery();
                        OleDbCommand sqlCommandTMP_EDS = new OleDbCommand(sqlDeleteTMP_EDS, connection);
                        sqlCommandTMP_EDS.ExecuteNonQuery();
                    }
                }
                catch (SqlException exsql)
                {
                    MainWindow.AlarmBox(exsql.Message);
                }
                try
                {
                    fileName = $"{waferText}.{zipfileName}.zip";
                    Unzip(fileName);
                    ShowValues();
                }
                catch (Exception ex)
                {
                    MainWindow.AlarmBox(ex.Message);
                }
            }
            else
            {
                MainWindow.AlarmBox("Barcode Mismatch !!!");
                ClearInvoiceValues();
                waferLot.Clear();
                UploadButton.IsEnabled = false;
            }
        }
        public static void Unzip(string FileName)
        {
            getplasmaStatus = null;
            int WFSEQ;
            int line = 0;
            int wfcountFail = 0;
            resultTmpData = null;
            string ConnectionString = ConfigurationManager.AppSettings["ConnetionStringDBRISTLSI"]!;
            string? ChangeDirectory = ConfigurationManager.AppSettings["ChangeDirectory"]!;
            string? FileToCopy = ConfigurationManager.AppSettings["CBOutputPath"] + FileName; /*Shared Folder*/
            string? NewCopyCB = ConfigurationManager.AppSettings["NewCopyCBPath"]!; /*for backup file before sending*/
            string? ProcessPath = ConfigurationManager.AppSettings["ProcessPath"]!;
            string? TMP_ORDERNO = null;
            string? tmpData1 = null;
            string? tmpData2 = null;
            sumchipCount = 0;
            sumwfCount = 0;
            tmpData = null;
            int countLine = 1;

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
                    return;
                }

                Directory.SetCurrentDirectory(ChangeDirectory);

                try
                {
                    ZipFile.ExtractToDirectory(NewCopyCB, ProcessPath);
                }
                catch
                {
                    MainWindow.AlarmBox($"Please check zip file in the {ChangeDirectory} location !!!");
                    return;
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
                    return;
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

                                            using (OleDbConnection connection = new OleDbConnection(ConnectionString))
                                            {
                                                connection.Open();
                                                string sqlInsert = "INSERT INTO WF_EDS (WFLOTNO,WFSEQ,CHIPCOUNT,INVOICENO) Values (?,?,?,?)";

                                                using (OleDbCommand commandInsert = new OleDbCommand(sqlInsert, connection))
                                                {
                                                    commandInsert.CommandType = CommandType.Text;
                                                    commandInsert.Parameters.AddWithValue("WFLOTNO", tmpwfLotno);
                                                    commandInsert.Parameters.AddWithValue("WFSEQ", WFSEQ);
                                                    commandInsert.Parameters.AddWithValue("CHIPCOUNT", waferChipcount);
                                                    commandInsert.Parameters.AddWithValue("INVOICENO", tmp_invoiceNo);
                                                    commandInsert.ExecuteNonQuery();
                                                }
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

                string ConnetionStringMapOnline = ConfigurationManager.AppSettings["ConnetionStringMapOnline"]!;
                string sqlSelectWFSEQ = "SELECT WFSEQ, CHIPCOUNT FROM WF_EDS ORDER BY WFSEQ";
                string sqlSelectCHIPMASTER = "SELECT * FROM CHIPMASTER WHERE CHIPMODELNAME = (?)";
                string sqlselectSample = "SELECT DeviceName, LotNo, FlagLastShipout FROM MapOnline.dbo.EDSFlow " +
                                         "WHERE DeviceName LIKE '%8' AND LotNo = @LotNo AND FlowName = 'output' AND FlagLastShipout = 1";

                using (OleDbConnection connection = new OleDbConnection(ConnectionString))
                {
                    connection.Open();
                    OleDbCommand sqlCommandQueryChip = new OleDbCommand(sqlSelectCHIPMASTER, connection);
                    sqlCommandQueryChip.CommandType = CommandType.Text;
                    sqlCommandQueryChip.Parameters.AddWithValue("?", tmp_ChipmodelName);

                    using (OleDbDataReader readerQueryChip = sqlCommandQueryChip.ExecuteReader())
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

                    OleDbCommand sqlCommandQueryWF = new OleDbCommand(sqlSelectWFSEQ, connection);

                    using (OleDbDataReader readerQueryWF = sqlCommandQueryWF.ExecuteReader())
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

                if (getlotStatus != "Sample")
                {
                    getfinseqno = null;
                    getfinseqno = SetSeq("");
                    string sqlInsertTMP_EDS = "INSERT INTO TMP_EDS (CHIPMODELNAME,WFLOTNO,WFCOUNT,CHIPCOUNT,INVOICENO,CASENO,OUTDIV,RECDIV,ORDERNO,PLASMA,WFDATA1,WFDATA2,SEQNO,WFCOUNT_FAIL) Values (?,?,?,?,?,?,?,?,?,?,?,?,?,?)";
                    using (OleDbConnection connection = new OleDbConnection(ConnectionString))
                    {
                        connection.Open();
                        OleDbCommand sqlCommandQuery = new OleDbCommand(sqlInsertTMP_EDS, connection);
                        sqlCommandQuery.CommandType = CommandType.Text;
                        sqlCommandQuery.Parameters.AddWithValue("?", tmp_ChipmodelName);
                        sqlCommandQuery.Parameters.AddWithValue("?", tmpwfLotno);
                        sqlCommandQuery.Parameters.AddWithValue("?", sumwfCount);
                        sqlCommandQuery.Parameters.AddWithValue("?", sumchipCount);
                        sqlCommandQuery.Parameters.AddWithValue("?", tmp_invoiceNo);
                        sqlCommandQuery.Parameters.AddWithValue("?", TMP_CASENO);
                        sqlCommandQuery.Parameters.AddWithValue("?", TMP_OUTDIV);
                        sqlCommandQuery.Parameters.AddWithValue("?", TMP_RECDIV);
                        sqlCommandQuery.Parameters.AddWithValue("?", TMP_ORDERNO);
                        sqlCommandQuery.Parameters.AddWithValue("?", getplasmaStatus);
                        sqlCommandQuery.Parameters.AddWithValue("?", tmpData1);
                        sqlCommandQuery.Parameters.AddWithValue("?", tmpData2);
                        sqlCommandQuery.Parameters.AddWithValue("?", getfinseqno);
                        sqlCommandQuery.Parameters.AddWithValue("?", wfcountFail);
                        sqlCommandQuery.ExecuteNonQuery();
                    }
                }
            }
            else
            {
                MainWindow.AlarmBox("Lot file zip not found, Please check !!!");
            }
        }

        private void ShowValues()
        {
            if (tmp_invoiceNo != null && getfinseqno != null || getlotStatus == "Sample")
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
            if (getlotStatus == "Sample")
            {
                invoiceNo.Text = tmp_invoiceNo;
                chipModel.Text = tmp_ChipmodelName;
                waferCount.Text = sumwfCount.ToString();
                chipCount.Text = sumchipCount.ToString();
                lotStatus.Text = getlotStatus;
                seqNo.Text = "No Data";
            }
            else
            {
                invoiceNo.Text = tmp_invoiceNo;
                chipModel.Text = tmp_ChipmodelName;
                waferCount.Text = sumwfCount.ToString();
                chipCount.Text = sumchipCount.ToString();
                lotStatus.Text = getlotStatus;
                seqNo.Text = getfinseqno;
            }
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
            waferLot.Clear();
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
            checkHasrowsNyuko = false;
            checkWFLOTNOnyuko = null;
            string ConnectionString = ConfigurationManager.AppSettings["ConnetionStringDBRISTLSI"]!;
            string sqlSelectChipnyuko = "SELECT WFLOTNO FROM CHIPNYUKO WHERE WFLOTNO = (?)";

            string sqlSelectCheckTMPEDS = "SELECT * FROM TMP_EDS";
            using (OleDbConnection connection = new OleDbConnection(ConnectionString))
            {
                connection.Open();
                using (OleDbCommand sqlCommandQuery = new OleDbCommand(sqlSelectCheckTMPEDS, connection))
                {
                    OleDbDataReader reader = sqlCommandQuery.ExecuteReader();
                    while (reader.Read())
                    {
                        checkWFLOTNOnyuko = reader.GetString(1);
                    }
                    reader.Close();
                }
            }

            using (OleDbConnection connection = new OleDbConnection(ConnectionString))
            {
                connection.Open();
                using (OleDbCommand sqlCommandQuery = new OleDbCommand(sqlSelectChipnyuko, connection))
                {
                    sqlCommandQuery.CommandType = CommandType.Text;
                    sqlCommandQuery.Parameters.AddWithValue("?", checkWFLOTNOnyuko);
                    OleDbDataReader reader = sqlCommandQuery.ExecuteReader();
                    checkHasrowsNyuko = reader.HasRows;
                    reader.Close();
                }
            }

            string sqlSelectTMPEDS = "SELECT * FROM TMP_EDS";
            using (OleDbConnection connection = new OleDbConnection(ConnectionString))
            {
                connection.Open();
                using (OleDbCommand sqlCommandQuery = new OleDbCommand(sqlSelectTMPEDS, connection))
                {
                    OleDbDataReader reader = sqlCommandQuery.ExecuteReader();
                    while (reader.Read())
                    {
                        string CHIPMODELNAME = reader.GetString(0);
                        string WFLOTNO = reader.GetString(1);
                        int WFCOUNT = (int)reader.GetDecimal(2);
                        int CHIPCOUNT = (int)reader.GetDecimal(3);
                        string INVOICENO = reader.GetString(4);
                        string CASENO = reader.GetString(5);
                        string OUTDIV = reader.GetString(6);
                        string RECDIV = reader.GetString(7);
                        string ORDERNO = reader.GetString(8);
                        string WFDATA1 = reader.GetString(9);
                        string WFDATA2 = reader.GetString(10);
                        string SEQNO = reader.GetString(11);
                        string PLASMA = reader.GetString(12);
                        int WFCOUNT_FAIL = (int)reader.GetDecimal(13);

                        string insertQuery = "INSERT INTO CHIPNYUKO (CHIPMODELNAME, MODELCODE1, MODELCODE2, WFLOTNO, SEQNO, ENO, RFSEQNO, OUTDIV, RECDIV, STOCKDATE, RETURNFLAG, WFCOUNT,CHIPCOUNT, " +
                                             "ORDERNO, HIGHREL, AGARIDATE, BUNKAN, RETURNCLASS, SLIPNO, SLIPNOEDA,STAFFNO, WFINPUT, TESTERNO, PROGRAMVER,RECYCLE, RINGNO, INVOICENO, HOLDFLAG,CASENO, " +
                                             "DIRECTCLASS, DELETEFLAG, WFDATA1, WFDATA2, TIMESTAMP, MOVE_ORDERNO, PLASMA,WFCOUNT_FAIL) " +
                                             "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";
                        using (OleDbCommand insertCommand = new OleDbCommand(insertQuery, connection))
                        {
                            insertCommand.CommandType = CommandType.Text;
                            insertCommand.Parameters.AddWithValue("?", CHIPMODELNAME);
                            insertCommand.Parameters.AddWithValue("?", "");
                            insertCommand.Parameters.AddWithValue("?", "");
                            insertCommand.Parameters.AddWithValue("?", WFLOTNO);
                            insertCommand.Parameters.AddWithValue("?", SEQNO);
                            insertCommand.Parameters.AddWithValue("?", "");
                            insertCommand.Parameters.AddWithValue("?", "");
                            insertCommand.Parameters.AddWithValue("?", OUTDIV);
                            insertCommand.Parameters.AddWithValue("?", RECDIV);
                            insertCommand.Parameters.AddWithValue("?", DateTime.Now.ToString("yyMMdd"));
                            insertCommand.Parameters.AddWithValue("?", "");
                            insertCommand.Parameters.AddWithValue("?", WFCOUNT);
                            insertCommand.Parameters.AddWithValue("?", CHIPCOUNT);
                            insertCommand.Parameters.AddWithValue("?", ORDERNO);
                            insertCommand.Parameters.AddWithValue("?", "");
                            insertCommand.Parameters.AddWithValue("?", DateTime.Now.ToString("yyMMdd"));
                            insertCommand.Parameters.AddWithValue("?", "");
                            insertCommand.Parameters.AddWithValue("?", "");
                            insertCommand.Parameters.AddWithValue("?", "");
                            insertCommand.Parameters.AddWithValue("?", "");
                            insertCommand.Parameters.AddWithValue("?", "001");
                            insertCommand.Parameters.AddWithValue("?", "");
                            insertCommand.Parameters.AddWithValue("?", "");
                            insertCommand.Parameters.AddWithValue("?", "");
                            insertCommand.Parameters.AddWithValue("?", "");
                            insertCommand.Parameters.AddWithValue("?", "");
                            insertCommand.Parameters.AddWithValue("?", INVOICENO);
                            insertCommand.Parameters.AddWithValue("?", "");
                            insertCommand.Parameters.AddWithValue("?", CASENO);
                            insertCommand.Parameters.AddWithValue("?", "");
                            insertCommand.Parameters.AddWithValue("?", "");
                            insertCommand.Parameters.AddWithValue("?", WFDATA1);
                            insertCommand.Parameters.AddWithValue("?", WFDATA2);
                            insertCommand.Parameters.AddWithValue("?", DateTime.Now.ToString());
                            insertCommand.Parameters.AddWithValue("?", "");
                            insertCommand.Parameters.AddWithValue("?", PLASMA);
                            insertCommand.Parameters.AddWithValue("?", WFCOUNT_FAIL);
                            insertCommand.ExecuteNonQuery();
                        }
                    }
                    reader.Close();
                }
            }
        }
        private static void UpdateChipzaiko()
        {
            checkHasrowsZaiko = false;
            checkWFLOTNOzaiko = null;
            string ConnectionString = ConfigurationManager.AppSettings["ConnetionStringDBRISTLSI"]!;
            string sqlSelectChipzaiko = "SELECT WFLOTNO FROM CHIPZAIKO WHERE WFLOTNO = (?)";

            string sqlSelectCheckTMPEDS = "SELECT * FROM TMP_EDS";
            using (OleDbConnection connection = new OleDbConnection(ConnectionString))
            {
                connection.Open();
                using (OleDbCommand sqlCommandQuery = new OleDbCommand(sqlSelectCheckTMPEDS, connection))
                {
                    OleDbDataReader reader = sqlCommandQuery.ExecuteReader();
                    while (reader.Read())
                    {
                        checkWFLOTNOzaiko = reader.GetString(1);
                    }
                    reader.Close();
                }
            }

            using (OleDbConnection connection = new OleDbConnection(ConnectionString))
            {
                connection.Open();
                using (OleDbCommand sqlCommandQuery = new OleDbCommand(sqlSelectChipzaiko, connection))
                {
                    sqlCommandQuery.CommandType = CommandType.Text;
                    sqlCommandQuery.Parameters.AddWithValue("?", checkWFLOTNOzaiko);
                    OleDbDataReader reader = sqlCommandQuery.ExecuteReader();
                    checkHasrowsZaiko = reader.HasRows;
                    reader.Close();
                }
            }

            string sqlSelectTMPEDS = "SELECT * FROM TMP_EDS";
            using (OleDbConnection connection = new OleDbConnection(ConnectionString))
            {
                connection.Open();
                using (OleDbCommand sqlCommandQuery = new OleDbCommand(sqlSelectTMPEDS, connection))
                {
                    OleDbDataReader reader = sqlCommandQuery.ExecuteReader();
                    while (reader.Read())
                    {
                        string CHIPMODELNAME = reader.GetString(0);
                        string WFLOTNO = reader.GetString(1);
                        int WFCOUNT = (int)reader.GetDecimal(2);
                        int CHIPCOUNT = (int)reader.GetDecimal(3);
                        string SEQNO = reader.GetString(11);
                        string insertQuery = "INSERT INTO CHIPZAIKO (CHIPMODELNAME, MODELCODE1, MODELCODE2, WFLOTNO, SEQNO, ENO, LOCATION, WFCOUNT, CHIPCOUNT, STOCKDATE, RETURNFLAG, REMAINFLAG, HOLDFLAG, STAFFNO, PREOUTFLAG, INVOICENO, PROCESSCODE, DELETEFLAG, TIMESTAMP) " +
                        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";
                        using (OleDbCommand insertCommand = new OleDbCommand(insertQuery, connection))
                        {
                            insertCommand.CommandType = CommandType.Text;
                            insertCommand.Parameters.AddWithValue("?", CHIPMODELNAME);
                            insertCommand.Parameters.AddWithValue("?", "");
                            insertCommand.Parameters.AddWithValue("?", "");
                            insertCommand.Parameters.AddWithValue("?", WFLOTNO);
                            insertCommand.Parameters.AddWithValue("?", SEQNO);
                            insertCommand.Parameters.AddWithValue("?", "");
                            insertCommand.Parameters.AddWithValue("?", "");
                            insertCommand.Parameters.AddWithValue("?", WFCOUNT);
                            insertCommand.Parameters.AddWithValue("?", CHIPCOUNT);
                            insertCommand.Parameters.AddWithValue("?", DateTime.Now.ToString("yyMMdd"));
                            insertCommand.Parameters.AddWithValue("?", "");
                            insertCommand.Parameters.AddWithValue("?", "");
                            insertCommand.Parameters.AddWithValue("?", "");
                            insertCommand.Parameters.AddWithValue("?", "001");
                            insertCommand.Parameters.AddWithValue("?", "");
                            insertCommand.Parameters.AddWithValue("?", "");
                            insertCommand.Parameters.AddWithValue("?", "TI970");
                            insertCommand.Parameters.AddWithValue("?", "");
                            insertCommand.Parameters.AddWithValue("?", DateTime.Now.ToString());
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
                    File.Move(FileToCopy, destinationFilePath + ".bak", true);
                    MainWindow.AlarmBox("Uploaded Successfully");
                }
                catch (Exception e)
                {
                    MainWindow.AlarmBox(e.Message);
                }
            }
            else
            {
                MainWindow.AlarmBox("Not found zip file in the CBOutput location !!!");
            }
        }
        private void UnzipSplitLot(string getLOTName, string getFileName)
        {
            string ExtractPath = ConfigurationManager.AppSettings.Get("ExtractPath") + getLOTName.TrimEnd();
            string CBOutputPath = ConfigurationManager.AppSettings.Get("CBOutputPath") + getFileName;
            if (Directory.Exists(ExtractPath))
            {
                ZipFile.ExtractToDirectory(CBOutputPath, ExtractPath, true);

                DirectoryInfo FileDatDirectory = new DirectoryInfo(ExtractPath);
                FileInfo[] DatFile = FileDatDirectory.GetFiles("W-NO-*.DAT");
                if (DatFile != null)
                {
                    foreach (FileInfo GetNumberDat in DatFile)
                    {
                        FileInfo[] GetNumberFile = DatFile;
                        int CountDatFile = GetNumberFile.Count();
                        string Lotdotdatpath = ExtractPath + @"\LOT.DAT";
                        using (BinaryWriter writer = new BinaryWriter(File.Open(Lotdotdatpath, FileMode.Open, FileAccess.ReadWrite)))
                        {
                            if (CountDatFile == 25)
                            {
                                int offset = 38; //position you want to start editing
                                byte[] set_data = new byte[] { 0x01,0x00,0x02,0x00,0x03,0x00,0x04,0x00,0x05,0x00,0x06,0x00,0x07,0x00,
                                                                   0x08,0x00,0x09,0x00,0x10,0x00,0x11,0x00,0x12,0x00,0x13,0x00,0x14,0x00,
                                                                   0x15,0x00,0x16,0x00,0x17,0x00,0x18,0x00,0x19,0x00,0x20,0x00,0x21,0x00,
                                                                   0x22,0x00,0x23,0x00,0x24,0x00,0x25,0x00 }; //new data
                                writer.Seek(offset, SeekOrigin.Begin); //move your cursor to the position
                                writer.Write(set_data); //write it     
                            }
                            else
                            {
                                int CheckNumberofWafer = 1;
                                foreach (FileInfo GetDatnumber in DatFile)
                                {
                                    string GetNumberWafer = GetDatnumber.Name.Substring(5, 2);
                                    int GetIntNumberDat = Convert.ToInt32(GetNumberWafer);
                                    byte GetByteNumberDat = Convert.ToByte(GetIntNumberDat);
                                    byte ZeroData = Convert.ToByte(0);
                                JumpForLoopSetData: //lopp for set data to correct position
                                    if (GetByteNumberDat == CheckNumberofWafer)
                                    {
                                        if (GetByteNumberDat == 1)
                                        {
                                            writer.Seek(38, SeekOrigin.Begin);
                                            writer.Write(1);
                                        }
                                        else if (GetByteNumberDat == 2)
                                        {
                                            writer.Seek(40, SeekOrigin.Begin);
                                            writer.Write(2);
                                        }
                                        else if (GetByteNumberDat == 3)
                                        {
                                            writer.Seek(42, SeekOrigin.Begin);
                                            writer.Write(3);
                                        }
                                        else if (GetByteNumberDat == 4)
                                        {
                                            writer.Seek(44, SeekOrigin.Begin);
                                            writer.Write(4);
                                        }
                                        else if (GetByteNumberDat == 5)
                                        {
                                            writer.Seek(46, SeekOrigin.Begin);
                                            writer.Write(5);
                                        }
                                        else if (GetByteNumberDat == 6)
                                        {
                                            writer.Seek(48, SeekOrigin.Begin);
                                            writer.Write(6);
                                        }
                                        else if (GetByteNumberDat == 7)
                                        {
                                            writer.Seek(50, SeekOrigin.Begin);
                                            writer.Write(7);
                                        }
                                        else if (GetByteNumberDat == 8)
                                        {
                                            writer.Seek(52, SeekOrigin.Begin);
                                            writer.Write(8);
                                        }
                                        else if (GetByteNumberDat == 9)
                                        {
                                            writer.Seek(54, SeekOrigin.Begin);
                                            writer.Write(9);
                                        }
                                        else if (GetByteNumberDat == 10)
                                        {
                                            writer.Seek(56, SeekOrigin.Begin);
                                            writer.Write(16);
                                        }
                                        else if (GetByteNumberDat == 11)
                                        {
                                            writer.Seek(58, SeekOrigin.Begin);
                                            writer.Write(17);
                                        }
                                        else if (GetByteNumberDat == 12)
                                        {
                                            writer.Seek(60, SeekOrigin.Begin);
                                            writer.Write(18);
                                        }
                                        else if (GetByteNumberDat == 13)
                                        {
                                            writer.Seek(62, SeekOrigin.Begin);
                                            writer.Write(19);
                                        }
                                        else if (GetByteNumberDat == 14)
                                        {
                                            writer.Seek(64, SeekOrigin.Begin);
                                            writer.Write(20);
                                        }
                                        else if (GetByteNumberDat == 15)
                                        {
                                            writer.Seek(66, SeekOrigin.Begin);
                                            writer.Write(21);
                                        }
                                        else if (GetByteNumberDat == 16)
                                        {
                                            writer.Seek(68, SeekOrigin.Begin);
                                            writer.Write(22);
                                        }
                                        else if (GetByteNumberDat == 17)
                                        {
                                            writer.Seek(70, SeekOrigin.Begin);
                                            writer.Write(23);
                                        }
                                        else if (GetByteNumberDat == 18)
                                        {
                                            writer.Seek(72, SeekOrigin.Begin);
                                            writer.Write(24);
                                        }
                                        else if (GetByteNumberDat == 19)
                                        {
                                            writer.Seek(74, SeekOrigin.Begin);
                                            writer.Write(25);
                                        }
                                        else if (GetByteNumberDat == 20)
                                        {
                                            writer.Seek(76, SeekOrigin.Begin);
                                            writer.Write(32);
                                        }
                                        else if (GetByteNumberDat == 21)
                                        {
                                            writer.Seek(78, SeekOrigin.Begin);
                                            writer.Write(33);
                                        }
                                        else if (GetByteNumberDat == 22)
                                        {
                                            writer.Seek(80, SeekOrigin.Begin);
                                            writer.Write(34);
                                        }
                                        else if (GetByteNumberDat == 23)
                                        {
                                            writer.Seek(82, SeekOrigin.Begin);
                                            writer.Write(35);
                                        }
                                        else if (GetByteNumberDat == 24)
                                        {
                                            writer.Seek(84, SeekOrigin.Begin);
                                            writer.Write(36);
                                        }
                                        else if (GetByteNumberDat == 25)
                                        {
                                            writer.Seek(86, SeekOrigin.Begin);
                                            writer.Write(37);
                                        }
                                        else
                                        {
                                            CheckNumberofWafer++;
                                        }
                                    }
                                    else
                                    {
                                        writer.Seek(0, SeekOrigin.Current); //move your cursor to the position
                                        writer.Write(ZeroData); //write data  
                                        CheckNumberofWafer++;
                                        goto JumpForLoopSetData;
                                    }
                                }
                            }
                            //Convert lot name Ascii to Hex Data 
                            string AsciiString = getLOTName.TrimEnd();
                            byte[] bytes = Encoding.Default.GetBytes(AsciiString);
                            int CharCount = AsciiString.Length;
                            for (int CharPosition = 1; CharPosition <= CharCount; CharPosition++)
                            {
                                writer.Seek(0, SeekOrigin.Begin); //move your cursor to the position
                                writer.Write(bytes); //write it     
                            }
                        }
                    }
                    string? CopytoPath = ConfigurationManager.AppSettings["CopytoPath"]!; /*Backup Folder*/
                    string destinationFilePath = System.IO.Path.Combine(CopytoPath, getFileName!);
                    File.Move(CBOutputPath, destinationFilePath + ".bak", true);
                    MainWindow.AlarmBox("Uploaded Successfully");
                }
                else
                {
                    MainWindow.AlarmBox("Not found a W-NO.DAT in the WaferMapping location !!!");
                }
            }
            else
            {
                MainWindow.AlarmBox("Not found a LOTNO folder in the WaferMapping location !!!");
            }
        }
    }
}