using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.IO.Compression;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using static ChipbankImport.ModalFD;
using System.Data.Common;
using System.Data;

namespace ChipbankImport
{
    public partial class ModalSampleLot : Window
    {
        public string? zipfileName { get; set; } //from submitButton_Click MainWindow
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
        public ModalSampleLot()
        {
            InitializeComponent();
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
            MessageBox.Show("wait code");
        }
        private void checkButton_Click(object sender, RoutedEventArgs e)
        {
            string waferText = waferLot.Text.ToString();
            if (waferText != "")
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
                string fileName = $"{waferText}.{zipfileName}.zip";
                Unzip(fileName);
                ShowValues();
            }
            else
            {
                MainWindow.AlarmBox("Barcode Mismatch !!!");
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

            if (File.Exists(ProcessPath + @"\ROBIN_L.CSV"))
            {
                string? ReadlineText = null;
                using (FileStream fileStream = new FileStream(ProcessPath + @"\ROBIN_L.CSV", FileMode.Open))
                {
                    using (StreamReader streamReader = new StreamReader(fileStream))
                    {
                        while ((ReadlineText = streamReader.ReadLine()) != null)
                        {
                            if (countLine == 1)
                            {
                                tmp_ChipmodelName = ReadlineText;
                            }
                            if (countLine == 7)
                            {
                                tmp_invoiceNo = ReadlineText.Trim();
                            }
                            countLine++;
                        }
                    }
                }
            }
            else
            {
                MainWindow.AlarmBox("Data not exist check ROBIN_L.CSV !!!");
            }
            if (File.Exists(ProcessPath + @"\CBinput.txt"))
            {
                using (FileStream fileStream = new FileStream(ProcessPath + @"\CBinput.txt", FileMode.Open))
                {
                    using (StreamReader streamReader = new StreamReader(fileStream))
                    {
                        string? ReadlineText = null;
                        while ((ReadlineText = streamReader.ReadLine()) != null)
                        {
                            line = line + 1;
                            if (line == 2)
                            {
                                tmpwfLotno = ReadlineText;
                            }
                            if (line >= 3)
                            {
                                WFSEQ = line - 2;
                                waferChipcount = ReadlineText;
                                if (waferChipcount != "")
                                {
                                    if (waferChipcount == "0")
                                    {
                                        wfcountFail = wfcountFail + 1;
                                    }
                                    else
                                    {
                                        if (WFSEQ.ToString().Length == 1)
                                        {
                                            tmpData = tmpData + "  " + WFSEQ.ToString();
                                        }
                                        else if (WFSEQ.ToString().Length == 2)
                                        {
                                            tmpData = tmpData + " " + WFSEQ.ToString();
                                        }
                                        else if (WFSEQ.ToString().Length == 3)
                                        {
                                            tmpData = tmpData + WFSEQ.ToString();
                                        }
                                        if (!string.IsNullOrEmpty(waferChipcount))
                                        {
                                            sumchipCount += int.Parse(waferChipcount);
                                            if (waferChipcount!.Length == 1)
                                            {
                                                tmpData = tmpData + "     " + waferChipcount!;
                                            }
                                            else if (waferChipcount!.Length == 2)
                                            {
                                                tmpData = tmpData + "    " + waferChipcount!;
                                            }
                                            else if (waferChipcount!.Length == 3)
                                            {
                                                tmpData = tmpData + "   " + waferChipcount!;
                                            }
                                            else if (waferChipcount!.Length == 4)
                                            {
                                                tmpData = tmpData + "  " + waferChipcount!;
                                            }
                                            else if (waferChipcount!.Length == 5)
                                            {
                                                tmpData = tmpData + " " + waferChipcount!;
                                            }
                                            else if (waferChipcount!.Length == 6)
                                            {
                                                tmpData = tmpData + waferChipcount;
                                            }
                                            tmpData1 = tmpData!;
                                            tmpData2 = tmpData!;
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
                        if (tmpData?.Length <= 180)
                        {
                            tmpData1 = tmpData?.Substring(0, Math.Min(tmpData.Length, 180));
                        }
                        else
                        {
                            tmpData1 = tmpData?.Substring(0, Math.Min(tmpData.Length, 180));
                            tmpData2 = tmpData?.Substring(180, Math.Max(0, Math.Min(tmpData.Length - 180, 180)));
                        }
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
                            getlotStatus = "Normal";
                        }
                    }
                }
                string TMP_CASENO = "      ";
                string TMP_OUTDIV = "QI830";
                string TMP_RECDIV = "TI970";
                string checkDigit = tmpwfLotno!.Substring(0, 1);
                if (Char.IsDigit(checkDigit[0]))
                {
                    TMP_ORDERNO = null;
                    TMP_ORDERNO = "E" + tmpwfLotno;
                }
                else
                {
                    TMP_ORDERNO = tmpwfLotno;
                }
                SETSEQ();
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
        }
        private void ShowValues()
        {
            if (tmp_invoiceNo != null && finseqno != null)
            {
                invoiceNo.Text = tmp_invoiceNo;
                chipModel.Text = tmp_ChipmodelName;
                waferCount.Text = sumwfCount.ToString();
                chipCount.Text = sumchipCount.ToString();
                lotStatus.Text = getlotStatus;
                seqNo.Text = finseqno;
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
                DataWafer dataWafer = new DataWafer();
                dataWafer.dataGridwafer.ItemsSource = dataWFSEQ.DefaultView;
                dataWafer.ShowDialog();
            }
            else
            {
                invoiceNo.Text = null;
                chipModel.Text = null;
                waferCount.Text = null;
                chipCount.Text = null;
                lotStatus.Text = null;
                seqNo.Text = null;
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
                    plasmaStatus.Text = null;
                }
            }
        }
    }
}
