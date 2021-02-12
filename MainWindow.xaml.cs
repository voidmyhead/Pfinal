using System;
using System.Windows;
using System.Data;
using System.Data.SqlClient; 
using Microsoft.Office.Interop.Excel;


namespace Pfinal
{
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Score_Insert_Click(object sender, RoutedEventArgs e)//성적입력 버튼
        {
            SqlConnection Conn = new SqlConnection();
            Conn.ConnectionString = "server=.\\SQLEXPRESS; database= student; user id = sa; pwd = 412563;";
            Conn.Open();

            //초기 성적입력시 누락된 값이 있다면 누락된 항목이 있다고 알리고 성적 입력을 진행하지 않는다.
            if (txtName.Text == "" || txtEnglishScore.Text == "" || txtMathScore.Text == "")
            {
                MessageBox.Show("누락된 항목이 있어 성적을 입력할 수 없습니다.");
                return; 
            }
            else//DB에 입력될 성적들을 메인윈도우의 레이블에 값들을 계산해서 넣는다.
            {
                Lb_SumScore.Content = (int.Parse(txtEnglishScore.Text) + int.Parse(txtMathScore.Text));
                Lb_AverageScore.Content = (Double.Parse(txtEnglishScore.Text) + Double.Parse(txtMathScore.Text)) / 2;
            }

            string Insert_query;
            Insert_query = "INSERT INTO tblStudent VALUES('"+txtName.Text+"',"+txtEnglishScore.Text+","+txtMathScore.Text+","+Lb_AverageScore.Content+","+Lb_SumScore.Content+")";

            SqlCommand comm = new SqlCommand(Insert_query, Conn);
            comm.ExecuteNonQuery();

            Conn.Close();
           
            MessageBox.Show("성공적으로 입력했습니다.");
        }

        private void Score_Print_Click(object sender, RoutedEventArgs e)//성적출력 버튼
        {
            SqlConnection Conn = new SqlConnection();
            Conn.ConnectionString = "server=.\\SQLEXPRESS; database = student; user id = sa; pwd = 412563;";
            Conn.Open();
            SqlCommand Command = new SqlCommand("SELECT * FROM tblStudent", Conn);
            SqlDataReader Reader = Command.ExecuteReader();

            if (Reader.Read() == false )//DB에 저장되어 있는 정보가 없다면 DB를 닫는다.
            {
                MessageBox.Show("입력되어있는 학생이 없습니다. 학생의 성적 정보를 먼저 입력해주세요.");
                Reader.Close();
                Conn.Close();
            }
            else //DB에 저장되어있는 정보가 있다면 Print_Score 창을 새로 띄워 DB에 저장되어 있는 학생들의 성적정보를 보여준다.
            {
                Conn.Close();
                Pfinal.Print_Score print_Score = new Pfinal.Print_Score();
                print_Score.Show();
            }
 
        }

        private void Score_Delete_Click(object sender, RoutedEventArgs e)//성적삭제 버튼
        {
            //성적 삭제 버튼은 학생의 이름을 키값으로 활용하여 진행하기 떄문에 삭제하고자 하는 학생의 이름이 없다면 더이상 진행하지 않는다.
            //미리 DB내용을 읽고 DB안에 학생의 정보가 없다면 해당 사유로 인해 더이상 진행하지 않는다고 알리며 더이상 진행하지 않는다.
            SqlConnection Conn = new SqlConnection();
            Conn.ConnectionString = "server=.\\SQLEXPRESS; database = student; user id = sa; pwd = 412563;";
            Conn.Open();
            SqlCommand Command3 = new SqlCommand("SELECT * FROM tblStudent WHERE Name = '" + txtName.Text + "'", Conn);
            SqlDataReader reader = Command3.ExecuteReader();

            if (txtName.Text == "")
            {
                MessageBox.Show("성적을 삭제할 학생의 이름이 누락되었습니다.");
                reader.Close();
                Conn.Close();
            }
            else if (reader.Read() == false)
            {
                MessageBox.Show($"{txtName.Text} 학생은 입력되어 있지않아 성적을 삭제할 수 없습니다.");
                reader.Close();
                Conn.Close();
            }
            else //정상적으로 삭제가 완료될경우, 메인 윈도우창에서 보이는 모든 값들은 빈칸으로 보이도록 초기화.
            {
                reader.Close();
                string delete_query;
                delete_query = "delete from tblStudent where Name = '"+txtName.Text+"'";
                SqlCommand Command = new SqlCommand(delete_query, Conn);
                Command.ExecuteNonQuery();
                Conn.Close();
                MessageBox.Show($"{txtName.Text} 학생의 정보가 성공적으로 삭제되었습니다.");

                txtName.Text = "";
                txtEnglishScore.Text = "";
                txtMathScore.Text = "";
                Lb_SumScore.Content = "";
                Lb_AverageScore.Content = "";
            }   
        }

        private void Score_Change_Click(object sender, RoutedEventArgs e)//성적변경 버튼
        {
            //성적 변경 버튼은 학생의 이름을 키값으로 활용하여 진행하기 떄문에 변경하고자 하는 학생의 이름이 없다면 더이상 진행하지 않는다.
            //성적변경을 하고자 한다면 DB에 입력되어 있는 학생들중 성적을 변경하고자 하는 학생의 이름과 변경하고자 하는 성적의 값을 최소 1개 이상 입력할것을 요구한다.
            int chg_count = 0;
            SqlConnection Conn = new SqlConnection();
            Conn.ConnectionString = "server=.\\SQLEXPRESS; database = student; user id = sa; pwd = 412563;";
            Conn.Open();
            SqlCommand Command3 = new SqlCommand("SELECT * FROM tblStudent WHERE Name = '" + txtName.Text + "'", Conn);
            SqlDataReader finder = Command3.ExecuteReader();
            string change_query;
            string change_E_query;
            string change_M_query;
            if (txtName.Text == "")
            {
                MessageBox.Show("성적을 변경할 학생의 이름이 누락되었습니다.");
                Conn.Close();
            }
            else if (finder.Read() == false)
            {
                MessageBox.Show($"{txtName.Text} 학생은 입력되어 있지않아 성적을 변경할 수 없습니다.");
            }
            else if (txtEnglishScore.Text == "" && txtMathScore.Text == "")
            {
                MessageBox.Show($"{txtName.Text} 학생의 변경하고자 하는 과목의 성적을 입력하세요");
                Conn.Close();
            }
            else
            {
                // 성적을 둘중에 하나만 입력했을 경우의 동작

                if (txtEnglishScore.Text == "")
                {//수학점수만 변경 하는 경우
                    finder.Close();
                    int M_score = int.Parse(txtMathScore.Text);
                    change_query = "update tblStudent set MathScore = " + txtMathScore.Text + "where Name = '" + txtName.Text + "'";
                    SqlCommand Command = new SqlCommand(change_query, Conn);
                    Command.ExecuteNonQuery();

                    MessageBox.Show($"{txtName.Text} 학생의 수학점수가 성공적으로 변경되었습니다.");
                    chg_count = chg_count + 1;
                }
                else if (txtMathScore.Text == "")
                {//영어점수만 변경하는 경우
                    finder.Close();
                    int E_score = int.Parse(txtEnglishScore.Text);
                    change_query = "update tblStudent set EnglishScore = " + E_score + "where Name = '" + txtName.Text + "'";
                    SqlCommand Command = new SqlCommand(change_query, Conn);
                    Command.ExecuteNonQuery();

                    MessageBox.Show($"{txtName.Text} 학생의 영어점수가 성공적으로 변경되었습니다.");
                    chg_count = chg_count + 1;
                }
                else if (txtEnglishScore.Text != "" && txtMathScore.Text != "")
                {
                    //학생의 모든 성적을 변경하는 경우의 동작
                    finder.Close();
                    int E_score = int.Parse(txtEnglishScore.Text);
                    int M_score = int.Parse(txtMathScore.Text);
                    change_E_query = "update tblStudent set EnglishScore = " + E_score + "where Name = '" + txtName.Text + "'";
                    change_M_query = "update tblStudent set MathScore = " + M_score + "where Name = '" + txtName.Text + "'"; ;

                    SqlCommand Command_E = new SqlCommand(change_E_query, Conn);
                    Command_E.ExecuteNonQuery();

                    SqlCommand Command_M = new SqlCommand(change_M_query, Conn);
                    Command_M.ExecuteNonQuery();

                    MessageBox.Show($"{txtName.Text} 학생의 모든 성적이 성공적으로 변경되었습니다.");
                    chg_count = chg_count + 1;
                }
            }
            //성적 변경후 DB에 점수합계와 평균점수 저장, 레이블에 변경된 학생의 점수대로 평균과 합계 표기
            if (chg_count == 1)
            {
                finder.Close();
                SqlCommand Command2 = new SqlCommand("SELECT Name, EnglishScore, MathScore FROM tblStudent WHERE Name = '" + txtName.Text + "'", Conn);
                SqlDataReader Reader = Command2.ExecuteReader();
                //DB로부터 데이터들을 Read()를 이용하여 읽어들이고, array1 배열에 저장한다.
                //Reader[@]의 데이터는 바로 사용하기 힘들기 때문에 배열에 저장했다.
                Reader.Read();
                string[] array1 = new string[] { Reader[1].ToString(), Reader[2].ToString() };
                Reader.Close();
                //성적을 변경한 학생의 총점을 변경하여 Lb.SumScore의 Content를 변경하여 보여주고 해당 값은 DB에 변경된 Sum값으로 update 한다.
                int Sum = int.Parse(array1[0]) + int.Parse(array1[1]);
                Lb_SumScore.Content = Sum;
                string change_S_query = "update tblStudent set Sum = " + Sum + "where Name = '" + txtName.Text + "'";
                SqlCommand Command_S = new SqlCommand(change_S_query, Conn);
   
                Command_S.ExecuteNonQuery();
                //성적을 변경한 학생의 총점을 변경하여 Lb.Average의 Content를 변경하여 보여주고 해당 값은 DB에 변경된 Average값으로 update 한다.
                double Average = Sum / 2;
                Lb_AverageScore.Content = Average;
                string change_A_query = "update tblStudent set Average = " + Average + "where Name = '" + txtName.Text + "'";
                SqlCommand Command_A = new SqlCommand(change_A_query, Conn);

                Command_A.ExecuteNonQuery();

                Reader.Close();
                Conn.Close();

            }
        }

        private void Score_Chart_Click(object sender, RoutedEventArgs e)// 성적차트 버튼
        {

            SqlConnection Conn = new SqlConnection();
            Conn.ConnectionString = "server=.\\SQLEXPRESS; database = student; user id = sa; pwd = 412563;";
            Conn.Open();
            SqlCommand Command = new SqlCommand("SELECT * FROM tblStudent", Conn);
            SqlDataReader Reader = Command.ExecuteReader();

            if (Reader.Read() == false)//DB에 저장되어 있는 정보가 없다면 DB를 닫는다.
            {
                MessageBox.Show("입력되어있는 학생이 없습니다. 학생의 성적 정보를 먼저 입력해주세요.");
                Reader.Close();
                Conn.Close();
            }
            else//DB에 저장되어있는 정보가 있다면 Chart 창을 새로 띄워 DB에 저장되어 있는 학생들의 총합성적차트를 보여준다.
            {
                Pfinal.Chart print_Chart = new Pfinal.Chart();
                print_Chart.ShowDialog();
            }
        }

        private void Clear_Click(object sender, RoutedEventArgs e)//Clear 버튼
        {
            txtName.Text = "";
            txtEnglishScore.Text = "";
            txtMathScore.Text = "";
            Lb_SumScore.Content = "";
            Lb_AverageScore.Content = "";
        }

        private void Down_Load_Click(object sender, RoutedEventArgs e)//다운로드 버튼 : DB에 저장되어 있는 정보를 EXCEL 파일로 저장한다.
        {
            Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();//Application application = new Application();
            Workbook workbook = application.Workbooks.Add();

            Worksheet worksheet = workbook.Worksheets.Item[1];
            worksheet.Name = "학생성적";

            //엑셀에 저장할 데이터를 DB에서 가져와서 ds라는 DataSet에 저장한다.
            DataSet ds = new DataSet();

            SqlConnection Conn = new SqlConnection();
            Conn.ConnectionString = "server=.\\SQLEXPRESS; database = student; user id = sa; pwd = 412563;";
            Conn.Open();
            string Command = "SELECT * FROM tblStudent";
            SqlDataAdapter adapter = new SqlDataAdapter(Command, Conn);
            

            adapter.Fill(ds, "tblData");
            string Command_s = "SELECT COUNT (*) FROM tblStudent";
            SqlCommand cnt = new SqlCommand(Command_s, Conn);
            int totalCount = Convert.ToInt32(cnt.ExecuteScalar());//전체 데이터의 행의 개수를 카운트하여 DB에 저장되어있는 학생들의 수를 파악하도록함

            Conn.Close();

            //엑셀파일에 각 열의 내용을 미리 넣어 사용하도록함
            worksheet.Cells[1, 1] = "이름";
            worksheet.Cells[1, 2] = "영어점수";
            worksheet.Cells[1, 3] = "수학점수";
            worksheet.Cells[1, 4] = "평균";
            worksheet.Cells[1, 5] = "총합";

            //반복문을 통해 ds 데이터셋의 내용을 엑셀파일에 넣도록 함
           for (int i = 0; i < 5; i++)
            {
                for (int j = 0;j<totalCount; j++)
                {
                    worksheet.Cells[j+2, i+1] = ds.Tables[0].Rows[j][i];
                }
            }

            MessageBox.Show("기본경로는 D드라이브 입니다. D드라이브에서 확인하세요.");

            //저장하려는 엑셀파일의 이름과 경로를 직접 넣음
            //이미 파일이 있다면 덮어쓰기로 저장하도록함.
            try 
            {
                workbook.SaveAs(Filename: @"D:\Student Score.xlsx");
            }
            catch (System.Runtime.InteropServices.COMException)//파일을 덮어쓴다고 할때, 이를 거부하면 직접 파일명과 경로를 설정할 수 있도록 함.
            {
                workbook.Close();
            }

        }
        private void END_Click(object sender, RoutedEventArgs e)//작업종료 버튼(프로그램 종료)
        {
            System.Diagnostics.Process.GetCurrentProcess().Kill();
        }

    }
}
