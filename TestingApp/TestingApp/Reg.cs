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
    public partial class Reg : Form
    {
        public Reg()
        {
            InitializeComponent();
            connect();
        }
        private string CS = "";
        private string Nosology = "SELECT IDNosology AS ID, Nosology AS Нозология FROM Nosology";
        private string Gruppa = " Select IDGruppa as ID, Naimenovanie as Наименование from Gruppa";
        private string Student = "SELECT        Student.IDStudent AS ID, Student.FullName AS ФИО, Student.[Position] AS Должность, Gruppa.IDGruppa AS Группа, Student.BirstDate AS [Дата рождения], Student.Age AS Возраст, " +
       " Student.RegDate AS[Дата регистрации], Nosology.Nosology AS Нозология FROM Nosology, Student, Gruppa Where Nosology.IDNosology = Student.IDNosology and Gruppa.IDGruppa=Student.IDGruppa";
      
        //метод для загрузки данных из БД в ComboBox
        private void Load_in_CB(string CS, string cmdT, ComboBox CB, string field1, string field2)
        {
           SqlDataAdapter Adapter = new SqlDataAdapter(cmdT, CS);
           DataSet ds = new DataSet();
            Adapter.Fill(ds, "Table");
          // привязка ComboBox к таблице БД
            CB.DataSource = ds.Tables["Table"];
            CB.DisplayMember = field1; //установка отображаемого в списке поля
            CB.ValueMember = field2; //установка ключевого поля

        }

        //функция для расчета возраста по дате рождения
        public static int CalculateAge(DateTime BirthDate)
	{
    int YearsPassed = DateTime.Now.Year - BirthDate.Year;
    if (DateTime.Now.Month<BirthDate.Month || (DateTime.Now.Month == BirthDate.Month && DateTime.Now.Day<BirthDate.Day))
    {
	        YearsPassed--;
	    }

	    return YearsPassed;
	}


private void Reg_Load(object sender, EventArgs e)
        {
           
            dateTimePicker1.Value = DateTime.Now;
            txtFullName.TabIndex = 0;
            CBSex.SelectedIndex = 0;
            //MessageBox.Show(CS);
            Load_in_CB(CS, Nosology, CBNosology, "Нозология", "ID");
            Load_in_CB(CS, Gruppa, CBGroup, "Наименование", "ID");
            bSave.Enabled = false;

        }
        SqlCommand myCommand;
        public  void connect()
        {
            Login frm = new Login();
            CS = frm.ConnectionString;
 
        }
    
        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            txtAge.Text = CalculateAge(dateTimePicker2.Value).ToString();
        }

        private void txtPass2_TextChanged(object sender, EventArgs e)
        {
            if (txtPass1.Text!="")
            {
                if (txtPass2.Text==txtPass1.Text)
                {
                    lblCorrectPass.Text = "Проверка пароля пройдена!";
                    lblCorrectPass.ForeColor = Color.Blue;
                    bSave.Enabled = true;

                }
                else
                {
                    lblCorrectPass.Text = "Проверка пароля не пройдена!";
                    lblCorrectPass.ForeColor = Color.Red;
                    bSave.Enabled = false;

                }
            }
        }

        private void bSave_Click(object sender, EventArgs e)
        {
            if ((txtFullName.Text == "") || (CBGroup.Text == "") || (CBNosology.Text == "")||(CBSex.Text==""))
            {
                MessageBox.Show("Заполните поля!");
            }
            else
            {
                //Проверка неповторяющегося пароля
                string sql = Student + " and Password=" + "'" +(txtPass1.Text).Trim() + "'";
                SqlDataAdapter A = new SqlDataAdapter(sql, CS);
                DataSet ds = new DataSet();
                A.Fill(ds, "Table");
                if (ds.Tables[0].Rows.Count>0)
                {
                    MessageBox.Show("Пользователь с таким паролем уже зарегистрирован!");
                }
                else
                {
                    //запись пользователя в БД
                    string AddStud = "INSERT INTO STUDENT (FullName,Position,IDGruppa,BirstDate,Age,RegDate,[Password],IDNosology)" +
                        " VALUES(" + "'" + txtFullName.Text + "'" + "," + "'" + txtPosition.Text + "'" + "," + "'" + CBGroup.SelectedValue+ "'" + ",'" +
                        dateTimePicker2.Value + "'," + txtAge.Text + ",'" + dateTimePicker1.Value + "'," +
                        "'" + txtPass1.Text + "'" + "," + CBNosology.SelectedValue + ")";
                    // MessageBox.Show(AddStud);
                    SqlConnection conn = new SqlConnection(CS);
                    conn.Open();
                    SqlCommand myCommand = new SqlCommand(AddStud, conn);
                    myCommand.CommandText = AddStud;
                    myCommand.ExecuteNonQuery();
                    conn.Close();
                    MessageBox.Show("Пользователь успешно зарегистрирован!");


                }
            }
        }
    }
}
