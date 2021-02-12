using System.Collections.Generic;
using System.Windows;
using System.Data.SqlClient;


namespace Pfinal
{
    /// <summary>
    /// Chart.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Chart : Window
    {
        public Chart()
        {
            InitializeComponent();

            SqlConnection Conn = new SqlConnection();
            Conn.ConnectionString = "server=.\\SQLEXPRESS; database = student; user id = sa; pwd = 412563;";
            Conn.Open();
            SqlCommand Command = new SqlCommand("SELECT * FROM tblStudent", Conn);
            SqlDataReader Reader = Command.ExecuteReader();

            List<KeyValuePair<string, int>> valueList = new List<KeyValuePair<string, int>>();

            while (Reader.Read()) {
                string[] array1 = new string[] { Reader[0].ToString(), Reader[4].ToString() };
                
                valueList.Add(new KeyValuePair<string, int>(array1[0], int.Parse(array1[1])));
            }

            xColumnChart.DataContext = valueList;

            Reader.Close();
            Conn.Close();

        }

    }
}
