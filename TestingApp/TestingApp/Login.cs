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
    public partial class Login : Form
    {
        public Login()
        {
            InitializeComponent();
            CS = "Data Source=IGOR\\MSSQLSERVER3;" +
           "Initial Catalog=Oprosnik;Integrated Security=True;Connect Timeout=30;" +
           "Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False";
        }
        private string dbPath;
        private string CS;
        private string Student = "SELECT Student.IDStudent AS ID, Student.FullName AS ФИО, Student.IDGruppa as Группа, Student.Position AS Должность, Student.BirstDate AS [Дата рождения], Student.Age AS Возраст, " +
    " Student.RegDate AS[Дата регистрации], Nosology.Nosology AS Нозология " +
            "FROM Nosology, Student, Gruppa Where Nosology.IDNosology = Student.IDNosology and Gruppa.IDGruppa=Student.IDGruppa and";
        public string DbPath{
            get
            {
                return dbPath;
            }  
            set
            {
                dbPath=value;
            }

            }
        public  string ConnectionString
        {
            get
            {
                return CS;
            }
            set
            {
                CS = value;
            }
        }
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Reg frm = new Reg();
            frm.ShowDialog();
        }

        private void Login_Load(object sender, EventArgs e)
        {
            CBUser.SelectedIndex = 0;
            txtPassw.TabIndex = 0;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (CBUser.SelectedIndex == 0)
            {

                string sql = Student + " Password=" + "'" + (txtPassw.Text).Trim() + "'";
                SqlDataAdapter Adapter = new SqlDataAdapter(sql, CS);
                DataSet ds = new DataSet();
                Adapter.Fill(ds, "Table");
                if (ds.Tables[0].Rows.Count == 0)
                {
                    MessageBox.Show("Пользователь с таким паролем не зарегистрирован!");
                }
                else
                {
                    Student frm = new Student();
                    this.Hide();
                    frm.lblFullName.Text = ds.Tables[0].Rows[0][1].ToString();
                    frm.lblID.Text = ds.Tables[0].Rows[0][0].ToString();
                    frm.Show();

                }
            }
            if (CBUser.SelectedIndex==1)
            {
                if (txtPassw.Text=="admin")
                {
                    Psycholog frm = new Psycholog();
                    frm.Show();
                    this.Hide();
                }
                else
                {
                    MessageBox.Show("Для пользователя " + CBUser.Text + " пароль неверный!");
                }
            }

        }
    }
}
