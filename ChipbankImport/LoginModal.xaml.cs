using System.Configuration;
using System.Data.SqlClient;
using System.Windows;
using System.Windows.Input;

namespace ChipbankImport
{
    public partial class LoginModal : Window
    {
        public string? getName { get; set; }
        public string? getID { get; set; }
        public LoginModal()
        {
            InitializeComponent();
            userTextBox.Focus();
        }

        private void loginButton_Click(object sender, RoutedEventArgs e)
        {
            UserLogin();
            Close();
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                DragMove();
            }
        }
        private void userTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                loginButton_Click(sender, e);
            }
        }

        private void ExitModalLogin_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void UserLogin()
        {
            string? empID = null;
            string? empName = null;
            string EmpCode = userTextBox.Text.ToString();
            string ConnetionStringerrreportsdb = ConfigurationManager.AppSettings["ConnetionStringerrreportsdb"]!;
            using (SqlConnection connection = new SqlConnection(ConnetionStringerrreportsdb))
            {
                connection.Open();
                string sqlSelectUser = "SELECT user_name, authority, full_name FROM Authority_table WHERE user_name = @user_name";
                SqlCommand sqlCommandQueryUser = new SqlCommand(sqlSelectUser, connection);
                sqlCommandQueryUser.Parameters.AddWithValue("@user_name", EmpCode);
                using (SqlDataReader readerQueryUser = sqlCommandQueryUser.ExecuteReader())
                {
                    if (readerQueryUser.HasRows)
                    {
                        while (readerQueryUser.Read())
                        {
                            empID = readerQueryUser.GetString(0);
                            empName = readerQueryUser.GetString(2);
                        }

                        Close();
                        SpecialModal specialModal = new SpecialModal();
                        specialModal.getID = empID;
                        specialModal.getName = empName;
                        specialModal.ShowDialog();
                    }
                    else
                    {
                        MainWindow.AlarmBox("User not register !!!");
                    }
                }
            }
        }
    }
}
