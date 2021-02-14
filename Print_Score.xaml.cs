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
                //strRecord를 해당 포맷으로 출력
                strRecord = string.Format("이름 : {0} 영어 : {1} 수학 : {2} 평균 : {3} 총점 : {4}", Reader[0], Reader[1], Reader[2], Reader[3], Reader[4]);
                Score_Viewer.Items.Add(strRecord);
            }
            Reader.Close();
            Conn.Close();
        }

        private void Re_Button_Click(object sender, RoutedEventArgs e) // 새로고침 버튼
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
                //지정된 형식에 맞는 문자열로 변환
                strRecord = string.Format("이름 : {0} 영어 : {1} 수학 : {2} 평균 : {3} 총점 : {4}", Reader[0], Reader[1], Reader[2], Reader[3], Reader[4]);
                Score_Viewer.Items.Add(strRecord);
            }

            Reader.Close();
            Conn.Close();
        }

        private void S_Button_Click(object sender, RoutedEventArgs e)//성적순 정렬
        {
            Score_Viewer.Items.Clear();
            SqlConnection Conn = new SqlConnection();
            Conn.ConnectionString = "server=.\\SQLEXPRESS; database = student; user id = sa; pwd = 412563;";
            Conn.Open();
            SqlCommand Command = new SqlCommand("SELECT * FROM tblStudent ORDER BY Sum DESC", Conn); //성적을 내림차순 정렬
            SqlDataReader Reader = Command.ExecuteReader();
            string strRecord;
            while (Reader.Read())
            {
                //지정된 형식에 맞는 문자열로 변환
                strRecord = string.Format("이름 : {0} 영어 : {1} 수학 : {2} 평균 : {3} 총점 : {4}", Reader[0], Reader[1], Reader[2], Reader[3], Reader[4]);
                Score_Viewer.Items.Add(strRecord);
            }

            Reader.Close();
            Conn.Close();
        }
    }
}
