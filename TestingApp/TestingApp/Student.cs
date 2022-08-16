using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace TestingApp
{
    public partial class Student : Form
    {
        string[] Quest = new string[30];
        int sh1, sh2, sh3, sh4, sh5;
        string answer;
        int i, n;

        public Student()
        {
            InitializeComponent();
            connect();
        }

        string CS = "";
        string Result = "SELECT Result.IDTesting AS [№тестирования], Testing.TestingDate AS Дата, Result.NQuest AS [№ вопроса], Result.ReactionTime AS [Время реакции], Result.Answer AS Ответ "+
       " FROM Result INNER JOIN "+
       " Testing ON Result.IDTesting = Testing.IDTesting";

        private void toolStripButton1_Click(object sender, EventArgs e)
        {

        }

        private void Student_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }

        private void Student_Load(object sender, EventArgs e)
        {
            lblDate.Text = DateTime.Now.ToShortDateString();
            lblTime.Text = DateTime.Now.ToLongTimeString();
            timer2.Enabled = false;
            timer3.Enabled = false;
            bStart.Enabled = true;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            lblTime.Text = DateTime.Now.ToLongTimeString();
            timer1.Start();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 1;
            txtDatetest.Text = lblDate.Text;
            NowTime= DateTime.Now;
            txtTimeStartTest.Text = NowTime.ToLongTimeString();
            timer2.Enabled = true;

        }
        DateTime NowTime;
        private void timer2_Tick(object sender, EventArgs e)
        {
            TimeSpan TikTime;
            TikTime = DateTime.Now - NowTime;
            txtTimeTest.Text = TikTime.ToString("hh\\:mm\\:ss");
            timer2.Start();


        }
        public void connect()
        {
            Login frm = new Login();
            CS = frm.ConnectionString;

        }
        int IDTesting = 0;
        private void button5_Click(object sender, EventArgs e)
        {
            sh1 = 0; sh2 = 0; sh3 = 0; sh4 = 0; sh5 = 0;
            //richTextBox2.Text = "Нажмите далее чтобы начать тестирование!";
            bStart.Enabled = false;
            NowTimeTest = DateTime.Now;
            timer3.Enabled = true;
            //Запись в БД
            string Testing = "INSERT INTO TESTING (IDStudent,TestingDate,TestingTime,Status) VALUES(" +
                lblID.Text + ", '" + txtDatetest.Text + "'," + "'" + DateTime.Now.ToLongTimeString() + "'" +
                "," + "'" + "не обработан" + "'" + ")";
            // MessageBox.Show(AddTesting);
            SqlConnection conn = new SqlConnection(CS);
            conn.Open();
            SqlCommand myCommand = new SqlCommand(Testing, conn);
            myCommand.CommandText = Testing;
            myCommand.ExecuteNonQuery();
            conn.Close();
            //возврат ID тестирования
            string max = "select MAX(IDTesting) from Testing";
            SqlDataAdapter A = new SqlDataAdapter(max, CS);
            DataSet ds = new DataSet();
            A.Fill(ds, "Table");
            IDTesting =int.Parse( ds.Tables[0].Rows[0][0].ToString());
            radioButton1.Visible = false;
            radioButton2.Visible = false;
            radioButton3.Visible = false;
            radioButton4.Visible = false;
            radioButton5.Visible = false;
            radioButton6.Visible = false;
            i = 0;
            n = 1;
        }

        private void bNext_Click(object sender, EventArgs e)
        {
            radioButton1.Visible = true;
            radioButton2.Visible = true;
            radioButton3.Visible = true;
            radioButton4.Visible = true;
            radioButton5.Visible = true;
            radioButton6.Visible = true;

            if (radioButton1.Checked == true)
            {
                if (i == 1 || i == 2 || i == 4 || i == 17 || i == 19 || i == 25)
                {
                    sh1 = sh1 + (-3);
                }
                else if (i == 3 || i == 7 || i == 8 || i == 10 || i == 18 || i == 30)
                {
                    sh2 = sh2 + (-3);    
                }
                else if (i == 5 || i == 6 || i == 13 || i == 14 || i == 16 || i == 22)
                {
                    sh3 = sh3 + (-3);
                }
                else if (i == 9 || i == 11 || i == 20 || i == 21 || i == 23 || i == 28)
                {
                    sh4 = sh4 + (-3);
                }
                else if (i == 12 || i == 15 || i == 24 || i == 26 || i == 27 || i == 29)
                {
                    sh5 = sh5 + (-3);
                }
                answer = "Полностью не согласен";
            }
            else if (radioButton2.Checked == true)
            {
                if (i == 1 || i == 2 || i == 4 || i == 17 || i == 19 || i == 25)
                {
                    sh1 = sh1 + (-2);
                }
                else if (i == 3 || i == 7 || i == 8 || i == 10 || i == 18 || i == 30)
                {
                    sh2 = sh2 + (-2);
                }
                else if (i == 5 || i == 6 || i == 13 || i == 14 || i == 16 || i == 22)
                {
                    sh3 = sh3 + (-2);
                }
                else if (i == 9 || i == 11 || i == 20 || i == 21 || i == 23 || i == 28)
                {
                    sh4 = sh4 + (-2);
                }
                else if (i == 12 || i == 15 || i == 24 || i == 26 || i == 27 || i == 29)
                {
                    sh5 = sh5 + (-2);
                }
                answer = "В основном не согласен";
            }
            else if (radioButton3.Checked == true)
            {
                if (i == 1 || i == 2 || i == 4 || i == 17 || i == 19 || i == 25)
                {
                    sh1 = sh1 + (-1);
                }
                else if (i == 3 || i == 7 || i == 8 || i == 10 || i == 18 || i == 30)
                {
                    sh2 = sh2 + (-1);
                }
                else if (i == 5 || i == 6 || i == 13 || i == 14 || i == 16 || i == 22)
                {
                    sh3 = sh3 + (-1);
                }
                else if (i == 9 || i == 11 || i == 20 || i == 21 || i == 23 || i == 28)
                {
                    sh4 = sh4 + (-1);
                }
                else if (i == 12 || i == 15 || i == 24 || i == 26 || i == 27 || i == 29)
                {
                    sh5 = sh5 + (-1);
                }
                answer = "Отчасти не согласен";
            }
            else if (radioButton4.Checked == true)
            {
                if (i == 1 || i == 2 || i == 4 || i == 17 || i == 19 || i == 25)
                {
                    sh1 = sh1 + 1;
                }
                else if (i == 3 || i == 7 || i == 8 || i == 10 || i == 18 || i == 30)
                {
                    sh2 = sh2 + 1;
                }
                else if (i == 5 || i == 6 || i == 13 || i == 14 || i == 16 || i == 22)
                {
                    sh3 = sh3 + 1;
                }
                else if (i == 9 || i == 11 || i == 20 || i == 21 || i == 23 || i == 28)
                {
                    sh4 = sh4 + 1;
                }
                else if (i == 12 || i == 15 || i == 24 || i == 26 || i == 27 || i == 29)
                {
                    sh5 = sh5 + 1;
                }
                answer = "Отчасти согласен";
            }
            else if (radioButton5.Checked == true)
            {
                if (i == 1 || i == 2 || i == 4 || i == 17 || i == 19 || i == 25)
                {
                    sh1 = sh1 + 2;
                }
                else if (i == 3 || i == 7 || i == 8 || i == 10 || i == 18 || i == 30)
                {
                    sh2 = sh2 + 2;
                }
                else if (i == 5 || i == 6 || i == 13 || i == 14 || i == 16 || i == 22)
                {
                    sh3 = sh3 + 2;
                }
                else if (i == 9 || i == 11 || i == 20 || i == 21 || i == 23 || i == 28)
                {
                    sh4 = sh4 + 2;
                }
                else if (i == 12 || i == 15 || i == 24 || i == 26 || i == 27 || i == 29)
                {
                    sh5 = sh5 + 2;
                }
                answer = "В основном согласен";
            }
            else if (radioButton6.Checked == true)
            {
                if (i == 1 || i == 2 || i == 4 || i == 17 || i == 19 || i == 25)
                {
                    sh1 = sh1 + 3;
                }
                else if (i == 3 || i == 7 || i == 8 || i == 10 || i == 18 || i == 30)
                {
                    sh2 = sh2 + 3;
                }
                else if (i == 5 || i == 6 || i == 13 || i == 14 || i == 16 || i == 22)
                {
                    sh3 = sh3 + 3;
                }
                else if (i == 9 || i == 11 || i == 20 || i == 21 || i == 23 || i == 28)
                {
                    sh4 = sh4 + 3;
                }
                else if (i == 12 || i == 15 || i == 24 || i == 26 || i == 27 || i == 29)
                {
                    sh5 = sh5 + 3;
                }
                answer = "Полностью согласен";
            }

            if (i < 30)
            {
                txtNumZad.Text = n.ToString() + " / 30";
                txtNumZad2.Text = n.ToString();

                Quest[0] = "Для меня как отрицательные, так и положительные эмоции служат источником знания, как поступать в жизни.";
                Quest[1] = "Отрицательные эмоции помогают мне понять, что я должен изменить в моей жизни.";
                Quest[2] = "Я спокоен, когда испытываю давление со стороны.";
                Quest[3] = "Я способен наблюдать изменение своих чувств. ";
                Quest[4] = "Когда необходимо, я могу быть спокойным и сосредоточенным, чтобы действовать в соответствии с запросами жизни. ";
                Quest[5] = "Когда необходимо, я могу вызвать у себя широкий спектр положительных эмоций, такие как веселье, радость, внутренний подъем и юмор. ";
                Quest[6] = "Я слежу за тем, как я себя чувствую. ";
                Quest[7] = "После того как что-то расстроило меня, я могу легко совладать со своими чувствами. ";
                Quest[8] = "Я способен выслушивать проблемы других людей. ";
                Quest[9] = "Я не зацикливаюсь на отрицательных эмоциях. ";
                Quest[10] = "Я чувствителен к эмоциональным потребностям других. ";
                Quest[11] = "Я могу действовать успокаивающе на других людей. ";
                Quest[12] = "Я могу заставить себя снова и снова встать перед лицом препятствия. ";
                Quest[13] = "Я стараюсь подходить творчески к жизненным проблемам. ";
                Quest[14] = "Я адекватно реагирую на настроения, побуждения и желания других людей. ";
                Quest[15] = "Я могу легко входить в состояние спокойствия, готовности и сосредоточенности. ";
                Quest[16] = "Когда позволяет время, я обращаюсь к своим негативным чувствам и разбираюсь, в чем проблема. ";
                Quest[17] = "Я способен быстро успокоиться после неожиданного огорчения. ";
                Quest[18] = "Знание моих истинных чувств важно для поддержания «хорошей формы». ";
                Quest[19] = "Я хорошо понимаю эмоции других людей, даже если они не выражены открыто. ";
                Quest[20] = "Я хорошо могу распознавать эмоции по выражению лица. ";
                Quest[21] = "Я могу легко отбросить негативные чувства, когда необходимо действовать. ";
                Quest[22] = "Я хорошо улавливаю знаки в общении, которые указывают на то, в чем другие нуждаются. ";
                Quest[23] = "Люди считают меня хорошим знатоком переживаний других людей. ";
                Quest[24] = "Люди, осознающие свои истинные чувства, лучше управляют своей жизнью. ";
                Quest[25] = "Я способен улучшить настроение других людей. ";
                Quest[26] = "Со мной можно посоветоваться по вопросам отношений между людьми. ";
                Quest[27] = "Я хорошо настраиваюсь на эмоции других людей. ";
                Quest[28] = "Я помогаю другим использовать их побуждения для достижения личных целей. ";
                Quest[29] = "Я могу легко отключиться от переживания неприятностей.";

                richTextBox2.Text = Quest[i];


                string AddAnsver = "INSERT INTO Result (IDTesting,NQuest,ReactionTime,Answer) VALUES(" +
                   IDTesting + "," + i + "," + "'" + txtReactionTime.Text + "'" + "," +
                  "'" + answer + "'" + ")";
                SqlConnection conn = new SqlConnection(CS);
                conn.Open();
                SqlCommand myCommand = new SqlCommand(AddAnsver, conn);
                myCommand.CommandText = AddAnsver;
                myCommand.ExecuteNonQuery();
                conn.Close();
                i++;
                n++;


                txtReactionTime.Clear();
            timer3.Enabled = true;
            NowTimeTest = DateTime.Now;
            timer3.Start();

            //MessageBox.Show("sh1= " + sh1 + " sh2= " + sh2 + " sh3= " + sh3 + " sh4= " + sh4 + " sh5= " + sh5);

                
            }
            else
            {
                //Запись в БД
               
              //MessageBox.Show("Я вышел!");
              

                timer2.Stop();
                txtDateFinal.Text = txtDatetest.Text;
                txtStartTimeFinal.Text = txtTimeStartTest.Text;
                txtTimeTestFinal.Text = txtTimeTest.Text;
                txtCountZadFinal.Text = i.ToString();
                //вывод результатов текущего тестирования из БД

                string ResultTesting = Result + " where Result.IDTesting=" + IDTesting;
                    SqlDataAdapter A = new SqlDataAdapter(ResultTesting, CS);
                DataSet ds = new DataSet();
                A.Fill(ds, "Table");
                dgv1.DataSource = ds.Tables[0].DefaultView;
                tabControl1.SelectedIndex = 2;
                string TestingUp = "Update TESTING SET Status = 'Обработан', Res1 = " + sh1 + ", Res2 = " + sh2 + ", Res3 =" + sh3 + ", Res4 = " + sh4 + ", Res5 =" + sh5 + " Where IDTesting=" + IDTesting;
                SqlConnection conn = new SqlConnection(CS);
                conn.Open();
                SqlCommand myCommand = new SqlCommand(TestingUp, conn);
                myCommand.CommandText = TestingUp;
                myCommand.ExecuteNonQuery();
                conn.Close();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Close();
        }

        DateTime NowTimeTest;
        private void timer3_Tick(object sender, EventArgs e)
        {
            TimeSpan TikTime;
            TikTime = DateTime.Now - NowTimeTest;
            txtReactionTime.Text = TikTime.ToString("hh\\:mm\\:ss");
            timer3.Start();
        }

        private void txtAnsver1_TextChanged(object sender, EventArgs e)
        {
            timer3.Stop();
            //timer3.Enabled = false;
        }

        private void bExit_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 0;
        }
    }
}
