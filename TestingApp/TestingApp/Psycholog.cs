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
using Excel = Microsoft.Office.Interop.Excel;

namespace TestingApp
{
    public partial class Psycholog : Form
    {
        int zn1, zn2, zn3, zn4, zn5, zn6, zn7, zn8, zn9, zn10;
        string Zn1, Zn2, Zn3, Zn4, Zn5;
        string n1, n2, n3, n4, n5;
        public Psycholog()
        {
            InitializeComponent();
            connect();
            conn(Testing, CS, dgv1);
        }
        string CS = "";
        private string Student = "SELECT IDStudent AS ID, FullName AS [Полное имя] FROM Student";
        private string Gruppa = "SELECT IDGruppa AS ID, Naimenovanie AS Наименование FROM Gruppa";
        private string Nosology = "SELECT IDNosology AS ID, Nosology AS Нозология FROM Nosology";
        string Testing = "SELECT Testing.IDTesting AS [№], Testing.TestingDate AS Дата, Testing.TestingTime AS Время, Student.FullName AS [ФИО тестируемого], Testing.Status AS Статус, Res1 as [Эмоциональная осведомленность], Res2 as [Управления эмоциями], Res3 as [Самомотивация], Res4 as [Эмпатия], Res5 as [Распознавание эмоций других] " +
 " FROM(Student INNER JOIN Testing ON Student.IDStudent = Testing.IDStudent)";
        public void connect()
        {
            Login frm = new Login();
            CS = frm.ConnectionString;

        }

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

        public void conn(string cmdT,string ConnectionString,DataGridView dgv)
        {
            SqlDataAdapter A = new SqlDataAdapter(cmdT, ConnectionString);
            DataSet ds = new DataSet();
            A.Fill(ds, "Table");
            dgv.DataSource = ds.Tables[0].DefaultView;
        }
        private void Psycholog_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }

        private void Psycholog_Load(object sender, EventArgs e)
        {
            lblDate.Text = DateTime.Now.ToShortDateString();
            lblTime.Text = DateTime.Now.ToLongTimeString();
            Load_in_CB(CS, Nosology, CBNosologiyaSearch, "Нозология", "ID");
            Load_in_CB(CS, Student, cbResStud, "Полное имя", "ID");
            Load_in_CB(CS, Gruppa, cbTestGr, "Наименование", "ID");
            Load_in_CB(CS, Student, cbStaStud, "Полное имя", "ID");
            Load_in_CB(CS, Nosology, cbTestNos, "Нозология", "ID");

            dataGridView4.Columns.Add("Column0", "Шкала");
            dataGridView4.Columns.Add("Column1", "Балл");
            dataGridView4.Columns.Add("Column2", "Интерпретация");
            dataGridView4.Rows.Add("МП", "");
            dataGridView4.Rows.Add("МУ", "");
            dataGridView4.Rows.Add("ВП", "");
            dataGridView4.Rows.Add("ВУ", "");
            dataGridView4.Rows.Add("ВЭ", "");
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            lblTime.Text = DateTime.Now.ToLongTimeString();
            timer1.Start();
        }

        private void dgv1_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //string IDTesting = dgv1[0,dgv1.CurrentRow.Index].Value.ToString();
            //Result frm = new Result();
            //frm.txtIDTesting.Text = IDTesting;
            //frm.ShowDialog();
        }

        private void dgv1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void groupBox5_Enter(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void groupBox4_Enter(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked==true)
            {
                string Zap = "SELECT Testing.IDTesting AS [№], Testing.TestingDate AS Дата, Testing.TestingTime AS Время, Student.FullName AS [ФИО тестируемого], Testing.Status AS Статус, Res1 as [Эмоциональная осведомленность], Res2 as [Управления эмоциями], Res3 as [Самомотивация], Res4 as [Эмпатия], Res5 as [Распознавание эмоций других] " +
 " FROM(Student INNER JOIN Testing ON Student.IDStudent = Testing.IDStudent) Where Res1 < '19'";
                SqlDataAdapter A = new SqlDataAdapter(Zap, CS);
                DataSet ds = new DataSet();
                A.Fill(ds, "Table");
                dgv1.DataSource = ds.Tables[0].DefaultView;
            }
            else
            {
                dgv1.SelectAll();
                dgv1.ClearSelection();
                dgv1.Columns.Clear();

            }
        }

        private void CBNosologiyaSearch_SelectedIndexChanged(object sender, EventArgs e)
        {
            string SearchNos = "SELECT Testing.IDTesting AS [№], Testing.TestingDate AS Дата, Testing.TestingTime AS Время, Student.FullName AS [ФИО тестируемого], Testing.Status AS Статус, Res1 as [Эмоциональная осведомленность], Res2 as [Управления эмоциями], Res3 as [Самомотивация], Res4 as [Эмпатия], Res5 as [Распознавание эмоций других] " +
 " FROM(Student INNER JOIN Testing ON Student.IDStudent = Testing.IDStudent " +
 " JOIN Nosology ON Student.IDNosology = Nosology.IDNosology) Where Res1 < '19' and Student.IDNosology =" + CBNosologiyaSearch.SelectedValue;
            SqlDataAdapter A = new SqlDataAdapter(SearchNos, CS);
            DataSet ds = new DataSet();
            A.Fill(ds, "Table");
            dgv1.DataSource = ds.Tables[0].DefaultView;
        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            string DatTest = "SELECT Testing.IDTesting AS [№], Testing.TestingDate AS Дата, Testing.TestingTime AS Время, Student.FullName AS [ФИО тестируемого], Testing.Status AS Статус, Res1 as [Эмоциональная осведомленность], Res2 as [Управления эмоциями], Res3 as [Самомотивация], Res4 as [Эмпатия], Res5 as [Распознавание эмоций других] " +
 " FROM(Student INNER JOIN Testing ON Student.IDStudent = Testing.IDStudent) Where Res1 < '19' and Testing.TestingDate <='"+ dateTimePicker2.Value +"' and Testing.TestingDate >='"+ dateTimePicker1.Value+"'";
            SqlDataAdapter A = new SqlDataAdapter(DatTest, CS);
            DataSet ds = new DataSet();
            A.Fill(ds, "Table");
            dgv1.DataSource = ds.Tables[0].DefaultView;
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            string DatTest = "SELECT Testing.IDTesting AS [№], Testing.TestingDate AS Дата, Testing.TestingTime AS Время, Student.FullName AS [ФИО тестируемого], Testing.Status AS Статус, Res1 as [Эмоциональная осведомленность], Res2 as [Управления эмоциями], Res3 as [Самомотивация], Res4 as [Эмпатия], Res5 as [Распознавание эмоций других] " +
 " FROM(Student INNER JOIN Testing ON Student.IDStudent = Testing.IDStudent) Where Res1 < '19' and Testing.TestingDate <='" + dateTimePicker2.Value + "' and Testing.TestingDate >='" + dateTimePicker1.Value + "'";
            SqlDataAdapter A = new SqlDataAdapter(DatTest, CS);
            DataSet ds = new DataSet();
            A.Fill(ds, "Table");
            dgv1.DataSource = ds.Tables[0].DefaultView;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            string Obr = "SELECT Testing.IDTesting AS [№], Testing.TestingDate AS Дата, Testing.TestingTime AS Время, Student.FullName AS [ФИО тестируемого], Testing.Status AS Статус, Res1 as [Эмоциональная осведомленность], Res2 as [Управления эмоциями], Res3 as [Самомотивация], Res4 as [Эмпатия], Res5 as [Распознавание эмоций других] " +
 " FROM(Student INNER JOIN Testing ON Student.IDStudent = Testing.IDStudent) Where Res1 < '19' and Status = 'Обработан'";
            SqlDataAdapter A = new SqlDataAdapter(Obr, CS);
            DataSet ds = new DataSet();
            A.Fill(ds, "Table");
            dgv1.DataSource = ds.Tables[0].DefaultView;
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            string Neobr = "SELECT Testing.IDTesting AS [№], Testing.TestingDate AS Дата, Testing.TestingTime AS Время, Student.FullName AS [ФИО тестируемого], Testing.Status AS Статус, Res1 as [Эмоциональная осведомленность], Res2 as [Управления эмоциями], Res3 as [Самомотивация], Res4 as [Эмпатия], Res5 as [Распознавание эмоций других] " +
 " FROM(Student INNER JOIN Testing ON Student.IDStudent = Testing.IDStudent) Where Res1 < '19' and Status = 'не обработан'";
            SqlDataAdapter A = new SqlDataAdapter(Neobr, CS);
            DataSet ds = new DataSet();
            A.Fill(ds, "Table");
            dgv1.DataSource = ds.Tables[0].DefaultView;
        }

        private void label16_Click(object sender, EventArgs e)
        {

        }

        private void label18_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label19_Click(object sender, EventArgs e)
        {

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void tabControl2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void tabPage4_Click(object sender, EventArgs e)
        {

        }

        private void tabPage5_Click(object sender, EventArgs e)
        {

        }

        private void tabPage6_Click(object sender, EventArgs e)
        {

        }

        private void tabPage7_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void tabPage8_Click(object sender, EventArgs e)
        {

        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void dateTimePicker3_ValueChanged(object sender, EventArgs e)
        {

        }

        private void bStat_Click(object sender, EventArgs e)
        {
            string StatStud = "Select Result.NQuest as [№ вопроса], Result.ReactionTime as [Время реакции], Result.Answer as [Ответ] " + 
                "from Student, Testing, Result " + 
                "Where Student.IDStudent=Testing.IDStudent and Testing.IDTesting=Result.IDTesting and Student.IDStudent=" + cbStaStud.SelectedValue + " and Testing.TestingDate="+"'"+dateTimePicker3.Value+"'"+
                " and Testing.TestingTime >"+"'"+dateTimePicker13.Value+"'"+"";
            SqlDataAdapter A = new SqlDataAdapter(StatStud, CS);
            DataSet ds = new DataSet();
            A.Fill(ds, "Table");
            dataGridView1.DataSource = ds.Tables[0].DefaultView;
        }

        private void bOt_Click(object sender, EventArgs e)
        {
            string StatGrup = "Select Student.FullName as [Ф.И.О Студента], Testing.TestingDate as [Дата тестирования], Testing.TestingTime as [Время тестирования], " +
                "COUNT(Result.NQuest) as [Общее количество ответов], " +
                "CAST(DATEADD(ms, AVG(CAST(DateDiff(ms, '00:00:00', ISNULL(Result.ReactionTime, '00:00:00')) as bigint)), '00:00:00') as time) as [Среднее время реакции],  SUM(iif(Result.ReactionTime < '00:00:01', 1, 0)) as [Кл-во недостоверных ответов] " +
                "from Student, Testing, Result, Gruppa " +
                " Where GRuppa.IDGruppa=Student.IDGruppa and Student.IDStudent=Testing.IDStudent and Testing.IDTesting=Result.IDTesting and Gruppa.IDGruppa=" + cbTestGr.SelectedValue + " and Testing.TestingDate >=" + "'" + dateTimePicker4.Value + "'" +
                " and Testing.TestingDate <=" + "'" + dateTimePicker5.Value + "'" + " GROUP BY Student.FullName, Testing.TestingDate, Testing.TestingTime";
            SqlDataAdapter A = new SqlDataAdapter(StatGrup, CS);
            DataSet ds = new DataSet();
            A.Fill(ds, "Table");
            dataGridView2.DataSource = ds.Tables[0].DefaultView;
        }

        private void TFIOTest_TextChanged(object sender, EventArgs e)
        {
            string SearchFIO = "SELECT Testing.IDTesting AS [№], Testing.TestingDate AS Дата, Testing.TestingTime AS Время, Student.FullName AS [ФИО тестируемого], Testing.Status AS Статус, Res1 as [Эмоциональная осведомленность], Res2 as [Управления эмоциями], Res3 as [Самомотивация], Res4 as [Эмпатия], Res5 as [Распознавание эмоций других] " +
 " FROM(Student INNER JOIN Testing ON Student.IDStudent = Testing.IDStudent) Where Res1 < '19' and Student.FullName LIKE " + "'" + TFIOTest.Text + "%" + "'";
            SqlDataAdapter A = new SqlDataAdapter(SearchFIO, CS);
            DataSet ds = new DataSet();
            A.Fill(ds, "Table");
            dgv1.DataSource = ds.Tables[0].DefaultView;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string d10 = dateTimePicker11.Value.Day.ToString() + "." + dateTimePicker11.Value.Month.ToString() + "." +
               dateTimePicker11.Value.Year.ToString();
            string d11 = dateTimePicker12.Value.Day.ToString() + "." + dateTimePicker12.Value.Month.ToString() + "." +
                 dateTimePicker12.Value.Year.ToString();

            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            Microsoft.Office.Interop.Excel.Range ExcelCells;
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
            ExcelApp.Cells[1, 1] = "Сводный отчет о результатах тестирования по группе за период С " + d10 + " По " + d11;
            for (int i = 0; i < dataGridView6.ColumnCount; i++)
            {
                ExcelApp.Cells[2, i + 1] = dataGridView6.Columns[i].HeaderCell.Value;
            }
            for (int i = 0; i < dataGridView6.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridView6.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 3, j + 1] = dataGridView6.Rows[i].Cells[j].Value.ToString();

                }
            }
            int istr = dataGridView6.Rows.Count + 1;
            ExcelCells = ExcelApp.Range[ExcelWorkSheet.Columns[2], ExcelWorkSheet.Columns[12]];
            ExcelCells.EntireColumn.AutoFit();
            ExcelCells = ExcelApp.Range[ExcelWorkSheet.Columns[1], ExcelWorkSheet.Columns[12]];
            ExcelCells.HorizontalAlignment = Excel.Constants.xlLeft;
            ExcelCells = ExcelApp.Range[ExcelWorkSheet.Cells[2, 1], ExcelWorkSheet.Cells[istr, 12]];
            ExcelCells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            ExcelCells.Borders.Weight = Excel.XlBorderWeight.xlThin;
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string d1 = dateTimePicker3.Value.Day.ToString() + "." + dateTimePicker3.Value.Month.ToString() + "." +
                 dateTimePicker3.Value.Year.ToString();
            string d2 = dateTimePicker13.Value.Hour.ToString() + "." + dateTimePicker13.Value.Minute.ToString() + "." +
                 dateTimePicker13.Value.Second.ToString();

            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            Microsoft.Office.Interop.Excel.Range ExcelCells;
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
            ExcelApp.Cells[1, 1] = "Статистика тестирования испытуемого " + cbStaStud.Text + " дата тестирования " + d1 + " время " + d2;
            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                ExcelApp.Cells[2, i + 1] = dataGridView1.Columns[i].HeaderCell.Value;
            }
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 3, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();

                }
            }
            int istr = dataGridView1.Rows.Count + 1;
            ExcelCells = ExcelApp.Range[ExcelWorkSheet.Columns[2], ExcelWorkSheet.Columns[3]];
            ExcelCells.EntireColumn.AutoFit();
            ExcelCells = ExcelApp.Range[ExcelWorkSheet.Columns[1], ExcelWorkSheet.Columns[3]];
            ExcelCells.HorizontalAlignment = Excel.Constants.xlLeft;
            ExcelCells = ExcelApp.Range[ExcelWorkSheet.Cells[2, 1], ExcelWorkSheet.Cells[istr, 3]];
            ExcelCells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            ExcelCells.Borders.Weight = Excel.XlBorderWeight.xlThin;
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }

        private void bOtNos_Click(object sender, EventArgs e)
        {
            string StatNos = "Select Student.FullName as [Ф.И.О Студента], Testing.TestingDate as [Дата тестирования], Testing.TestingTime as [Время тестирования], " +
               "COUNT(Result.NQuest) as [Общее количество ответов], " +
               "CAST(DATEADD(ms, AVG(CAST(DateDiff(ms, '00:00:00', ISNULL(Result.ReactionTime, '00:00:00')) as bigint)), '00:00:00') as time) as [Среднее время реакции],  SUM(iif(Result.ReactionTime < '00:00:01', 1, 0)) as [Кл-во недостоверных ответов] " +
               "from Student, Testing, Result, Nosology " +
               " Where Nosology.IDNosology=Student.IDNosology and Student.IDStudent=Testing.IDStudent and Testing.IDTesting=Result.IDTesting and Student.IDNosology=" + cbTestNos.SelectedValue + " and Testing.TestingDate >=" + "'" + dateTimePicker6.Value + "'" +
               " and Testing.TestingDate <=" + "'" + dateTimePicker7.Value + "'" + " GROUP BY Student.FullName, Testing.TestingDate, Testing.TestingTime";
            SqlDataAdapter A = new SqlDataAdapter(StatNos, CS);
            DataSet ds = new DataSet();
            A.Fill(ds, "Table");
            dataGridView3.DataSource = ds.Tables[0].DefaultView;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string d3 = dateTimePicker4.Value.Day.ToString() + "." + dateTimePicker4.Value.Month.ToString() + "." +
                dateTimePicker4.Value.Year.ToString();
            string d4 = dateTimePicker5.Value.Day.ToString() + "." + dateTimePicker5.Value.Month.ToString() + "." +
                 dateTimePicker5.Value.Year.ToString();

            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            Microsoft.Office.Interop.Excel.Range ExcelCells;
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
            ExcelApp.Cells[1, 1] = "Отчет о тестировании студентов группы " + cbTestGr.Text + " за период с " + d3 + " по " + d4;
            for (int i = 0; i < dataGridView2.ColumnCount; i++)
            {
                ExcelApp.Cells[2, i + 1] = dataGridView2.Columns[i].HeaderCell.Value;
            }
            for (int i = 0; i < dataGridView2.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridView2.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 3, j + 1] = dataGridView2.Rows[i].Cells[j].Value.ToString();

                }
            }
            int istr = dataGridView2.Rows.Count + 1;
            ExcelCells = ExcelApp.Range[ExcelWorkSheet.Columns[2], ExcelWorkSheet.Columns[6]];
            ExcelCells.EntireColumn.AutoFit();
            ExcelCells = ExcelApp.Range[ExcelWorkSheet.Columns[1], ExcelWorkSheet.Columns[6]];
            ExcelCells.HorizontalAlignment = Excel.Constants.xlLeft;
            ExcelCells = ExcelApp.Range[ExcelWorkSheet.Cells[2, 1], ExcelWorkSheet.Cells[istr, 6]];
            ExcelCells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            ExcelCells.Borders.Weight = Excel.XlBorderWeight.xlThin;
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            string d5 = dateTimePicker6.Value.Day.ToString() + "." + dateTimePicker6.Value.Month.ToString() + "." +
                dateTimePicker6.Value.Year.ToString();
            string d6 = dateTimePicker7.Value.Day.ToString() + "." + dateTimePicker7.Value.Month.ToString() + "." +
                 dateTimePicker7.Value.Year.ToString();

            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            Microsoft.Office.Interop.Excel.Range ExcelCells;
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
            ExcelApp.Cells[1, 1] = "Отчет о тестировании испытуемых нозологии " + cbTestNos.Text + " за период с " + d5 + " по " + d6;
            for (int i = 0; i < dataGridView3.ColumnCount; i++)
            {
                ExcelApp.Cells[2, i + 1] = dataGridView3.Columns[i].HeaderCell.Value;
            }
            for (int i = 0; i < dataGridView3.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridView3.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 3, j + 1] = dataGridView3.Rows[i].Cells[j].Value.ToString();

                }
            }
            int istr = dataGridView3.Rows.Count + 1;
            ExcelCells = ExcelApp.Range[ExcelWorkSheet.Columns[2], ExcelWorkSheet.Columns[6]];
            ExcelCells.EntireColumn.AutoFit();
            ExcelCells = ExcelApp.Range[ExcelWorkSheet.Columns[1], ExcelWorkSheet.Columns[6]];
            ExcelCells.HorizontalAlignment = Excel.Constants.xlLeft;
            ExcelCells = ExcelApp.Range[ExcelWorkSheet.Cells[2, 1], ExcelWorkSheet.Cells[istr, 6]];
            ExcelCells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            ExcelCells.Borders.Weight = Excel.XlBorderWeight.xlThin;
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            string d7 = dateTimePicker8.Value.Day.ToString() + "." + dateTimePicker8.Value.Month.ToString() + "." +
               dateTimePicker8.Value.Year.ToString();


            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            Microsoft.Office.Interop.Excel.Range ExcelCells;
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
            ExcelApp.Cells[1, 1] = "Результаты тестирования испытуемого " + cbResStud.Text + " дата тестирования " + d7;
            for (int i = 0; i < dataGridView4.ColumnCount; i++)
            {
                ExcelApp.Cells[2, i + 1] = dataGridView4.Columns[i].HeaderCell.Value;
            }
            for (int i = 0; i < dataGridView4.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridView4.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 3, j + 1] = dataGridView4.Rows[i].Cells[j].Value.ToString();

                }
            }
            int istr = dataGridView4.Rows.Count + 1;
            ExcelCells = ExcelApp.Range[ExcelWorkSheet.Columns[2], ExcelWorkSheet.Columns[3]];
            ExcelCells.EntireColumn.AutoFit();
            ExcelCells = ExcelApp.Range[ExcelWorkSheet.Columns[1], ExcelWorkSheet.Columns[3]];
            ExcelCells.HorizontalAlignment = Excel.Constants.xlLeft;
            ExcelCells = ExcelApp.Range[ExcelWorkSheet.Cells[2, 1], ExcelWorkSheet.Cells[istr, 3]];
            ExcelCells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            ExcelCells.Borders.Weight = Excel.XlBorderWeight.xlThin;
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            string d8 = dateTimePicker9.Value.Day.ToString() + "." + dateTimePicker9.Value.Month.ToString() + "." +
               dateTimePicker9.Value.Year.ToString();
            string d9 = dateTimePicker10.Value.Day.ToString() + "." + dateTimePicker10.Value.Month.ToString() + "." +
                 dateTimePicker10.Value.Year.ToString();

            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            Microsoft.Office.Interop.Excel.Range ExcelCells;
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
            ExcelApp.Cells[1, 1] = "Сводный отчет о результатах тестирования по нозологии за период С " + d8 + " По " + d9;
            for (int i = 0; i < dataGridView5.ColumnCount; i++)
            {
                ExcelApp.Cells[2, i + 1] = dataGridView5.Columns[i].HeaderCell.Value;
            }
            for (int i = 0; i < dataGridView5.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridView5.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 3, j + 1] = dataGridView5.Rows[i].Cells[j].Value.ToString();

                }
            }
            int istr = dataGridView5.Rows.Count + 1;
            ExcelCells = ExcelApp.Range[ExcelWorkSheet.Columns[2], ExcelWorkSheet.Columns[12]];
            ExcelCells.EntireColumn.AutoFit();
            ExcelCells = ExcelApp.Range[ExcelWorkSheet.Columns[1], ExcelWorkSheet.Columns[12]];
            ExcelCells.HorizontalAlignment = Excel.Constants.xlLeft;
            ExcelCells = ExcelApp.Range[ExcelWorkSheet.Cells[2, 1], ExcelWorkSheet.Cells[istr, 12]];
            ExcelCells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            ExcelCells.Borders.Weight = Excel.XlBorderWeight.xlThin;
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }

        private void bResGr_Click(object sender, EventArgs e)
        {
            string OtchGR = "select Gruppa.Naimenovanie as [Наименование группы], " +
                "COUNT(Student.IDStudent) as [Кол - во], " +
                "AVG(Testing.Res1) as [Эмоциональная осведомленность], A" +
                "VG(Testing.Res2) as [Управление своими эмоциями], " +
                "AVG(Testing.Res3) as [Самомотивация], " +
                "AVG(Testing.Res4) as [Эмпатия], " +
                "AVG(Testing.Res5) as [Распознавание эмоций других людей] " +
" From Gruppa, Student, Testing " +
" Where Student.IDStudent = Testing.IDStudent and Gruppa.IDGruppa = Student.IDGruppa and Testing.TestingDate >=" + "'" + dateTimePicker11.Value + "'" +
               " and Testing.TestingDate <=" + "'" + dateTimePicker12.Value + "'" +
" Group by Gruppa.Naimenovanie ";
            SqlDataAdapter A = new SqlDataAdapter(OtchGR, CS);
            DataSet ds = new DataSet();
            A.Fill(ds, "Table");
            dataGridView6.DataSource = ds.Tables[0].DefaultView;
        }

        private void bResNos_Click(object sender, EventArgs e)
        {
            string OtchNz = "select Nosology.Nosology as [Наименование нозологии], " +
                "COUNT(Student.IDStudent) as [Кол - во], " +
                "AVG(Testing.Res1) as [Эмоциональная осведомленность], A" +
                "VG(Testing.Res2) as [Управление своими эмоциями], " +
                "AVG(Testing.Res3) as [Самомотивация], " +
                "AVG(Testing.Res4) as [Эмпатия], " +
                "AVG(Testing.Res5) as [Распознавание эмоций других людей] " +
" From Nosology, Student, Testing " +
" Where Student.IDStudent = Testing.IDStudent and Nosology.IDNosology=Student.IDNosology and Testing.TestingDate >=" + "'" + dateTimePicker9.Value + "'" +
               " and Testing.TestingDate <=" + "'" + dateTimePicker10.Value + "'" +
" Group by Nosology.Nosology ";
            SqlDataAdapter A = new SqlDataAdapter(OtchNz, CS);
            DataSet ds = new DataSet();
            A.Fill(ds, "Table");
            dataGridView5.DataSource = ds.Tables[0].DefaultView;
        }

        private void bRes_Click(object sender, EventArgs e)
        {
            string ResStud = "Select Res1, Res2, Res3, Res4, Res5 " +
               "from Student, Testing " +
               "Where Student.IDStudent=Testing.IDStudent and Student.IDStudent=" + cbResStud.SelectedValue + " and Testing.TestingDate=" + "'" + dateTimePicker8.Value + "'" + "";
            SqlDataAdapter A = new SqlDataAdapter(ResStud, CS);
            DataSet ds = new DataSet();
            A.Fill(ds, "Table");

            n1 = ds.Tables[0].Rows[0][0].ToString();
            n2 = ds.Tables[0].Rows[0][1].ToString();
            n3 = ds.Tables[0].Rows[0][2].ToString();
            n4 = ds.Tables[0].Rows[0][3].ToString();
            n5 = ds.Tables[0].Rows[0][4].ToString();

            int zn1 = Int32.Parse(n1);
            int zn2 = Int32.Parse(n2);
            int zn3 = Int32.Parse(n3);
            int zn4 = Int32.Parse(n4);
            int zn5 = Int32.Parse(n5);

            if (zn1 <= 7)
            {
                Zn1 = "Низкий";
            }
            else if (zn1 >= 8 && zn1 <= 13)
            {
                Zn1 = "Средний";
            }
            else 
            {
                Zn1 = "Высокий";
            }
          

            if (zn2 <= 7)
            {
                Zn2 = "Низкий";
            }
            else if (zn2 >= 8 && zn2 <= 13)
            {
                Zn2 = "Средний";
            }
            else 
            {
                Zn2 = "Высокий";
            }
            

            if (zn3 <= 7)
            {
                Zn3 = "Низкий";
            }
            else if (zn3 >= 8 && zn3 <= 13)
            {
                Zn3 = "Средний";
            }
            else
            {
                Zn3 = "Высокий";
            }
            

            if (zn4 <= 7)
            {
                Zn4 = "Низкий";
            }
            else if (zn4 >= 8 && zn4 <= 13)
            {
                Zn4 = "Средний";
            }
            else 
            {
                Zn4 = "Высокий";
            }
            

            if (zn5 <= 7)
            {
                Zn5 = "Низкий";
            }
            else if (zn5 >= 8 && zn5 <= 13)
            {
                Zn5 = "Средний";
            }
            else 
            {
                Zn5 = "Высокий";
            }
  

            dataGridView4[1, 0].Value = ds.Tables[0].Rows[0][0].ToString();
            dataGridView4[1, 1].Value = ds.Tables[0].Rows[0][1].ToString();
            dataGridView4[1, 2].Value = ds.Tables[0].Rows[0][2].ToString();
            dataGridView4[1, 3].Value = ds.Tables[0].Rows[0][3].ToString();
            dataGridView4[1, 4].Value = ds.Tables[0].Rows[0][4].ToString();

            dataGridView4[2, 0].Value = Zn1;
            dataGridView4[2, 1].Value = Zn2;
            dataGridView4[2, 2].Value = Zn3;
            dataGridView4[2, 3].Value = Zn4;
            dataGridView4[2, 4].Value = Zn5;
        }
    }
}
