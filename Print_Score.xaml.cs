using System.Windows;
using System.Data.SqlClient;

namespace Pfinal
{

    public partial class Print_Score : Window
    {
        public Print_Score()
        {
            InitializeComponent();
            SqlConnection Conn = new SqlConnection();
            Conn.ConnectionString = "server=.\\SQLEXPRESS; database = student; user id = sa; pwd = 412563;";
            Conn.Open();
            SqlCommand Command = new SqlCommand("SELECT * FROM tblStudent", Conn);
            SqlDataReader Reader = Command.ExecuteReader();
            string strRecord;
            while (Reader.Read())
            { 
                strRecord = string.Format("이름 : {0} 영어 : {1} 수학 : {2} 평균 : {3} 총점 : {4}", Reader[0], Reader[1], Reader[2], Reader[3], Reader[4]);
                Score_Viewer.Items.Add(strRecord);
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e) // 새로고침
        {
            Score_Viewer.Items.Clear();
            SqlConnection Conn = new SqlConnection();
            Conn.ConnectionString = "server=.\\SQLEXPRESS; database = student; user id = sa; pwd = 412563;";
            Conn.Open();
            SqlCommand Command = new SqlCommand("SELECT * FROM tblStudent", Conn);
            SqlDataReader Reader = Command.ExecuteReader();
            string strRecord;
            while (Reader.Read())
            {
                strRecord = string.Format("이름 : {0} 영어 : {1} 수학 : {2} 평균 : {3} 총점 : {4}", Reader[0], Reader[1], Reader[2], Reader[3], Reader[4]);
                Score_Viewer.Items.Add(strRecord);
            }

            Reader.Close();
            Conn.Close();
        }
    }
}
