using System.Configuration;
using System.Data.SqlClient;
using System.Windows;
using System.Windows.Input;

namespace ChipbankImport
{
    public partial class DataWafer : Window
    {
        public string? tmpwfLotno { get; set; }
        public int? sumwfCount { get; set; }
        public int? sumchipCount { get; set; }
        public bool checkwaferFail { get; set; }
        public DataWafer()
        {
            InitializeComponent();
        }

        private void ExitModalFD_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void closeButton_Click(object sender, RoutedEventArgs e)
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

        private void Window_ContentRendered(object sender, System.EventArgs e)
        {
            string ConnectionString = ConfigurationManager.AppSettings["ConnetionStringMapOnline"]!;
            string sqlselectCheckpcs = "SELECT InputLotNo, EndWaferPcs, EndChipPcs FROM EDSFlow WHERE InputLotNo = @tmpwfLotno AND FlowName = 'OUTPUT' AND FlagLastShipout = 1";
            using (SqlConnection connection = new SqlConnection(ConnectionString))
            {
                string? InputLotNo = null;
                int EndWaferPcs = 0;
                int EndChipPcs = 0;
                connection.Open();
                SqlCommand sqlCommandQuerypcs = new SqlCommand(sqlselectCheckpcs, connection);
                sqlCommandQuerypcs.Parameters.AddWithValue("@tmpwfLotno", tmpwfLotno);
                using (SqlDataReader readerQuerypcs = sqlCommandQuerypcs.ExecuteReader())
                {
                    if (readerQuerypcs.HasRows)
                    {
                        while (readerQuerypcs.Read())
                        {
                            InputLotNo = readerQuerypcs.GetString(0);
                            EndWaferPcs = readerQuerypcs.GetInt32(1);
                            EndChipPcs = readerQuerypcs.GetInt32(2);
                        }
                        bool LotEqual = (InputLotNo == tmpwfLotno);
                        bool WaferEqual = (EndWaferPcs == sumwfCount);
                        bool ChipEqual = (EndChipPcs == sumchipCount);
                        if (!LotEqual || !WaferEqual || !ChipEqual)
                        {
                            MainWindow.AlarmBox("Data is not correct please check !!!");
                            checkwaferFail = true;
                        }
                    }
                    else
                    {
                        MainWindow.AlarmBox("Data is not correct please check !!!");
                        checkwaferFail = true;
                    }
                }
            }
        }
    }
}
