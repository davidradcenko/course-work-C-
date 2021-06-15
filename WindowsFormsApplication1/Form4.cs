using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;

namespace WindowsFormsApplication1
{
    public partial class Form4 : Form
    {
        OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Control_za_studentami.mdb");
        
        public Form4()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form2 ee = new Form2();
            ee.Show();
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

       

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            Bitmap bm = new Bitmap(this.dataGridView1.Width, this.dataGridView1.Height);

            dataGridView1.DrawToBitmap(bm, new Rectangle(0, 0, this.dataGridView1.Width, this.dataGridView1.Height));
            e.Graphics.DrawImage(bm, 0, 0);
        }
        private void button1_Click(object sender, EventArgs e)
        {
            printDocument1.Print();
        }

        private void button4_Click(object sender, EventArgs e)
        {

            if (radioButton1.Checked && !(comboBox1.Text == ""))
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                string s = comboBox1.Text;

                if ("Все" == comboBox1.Text)
                {
                    s = "СП";
                }

                cmd.CommandText = "Select   Студенты.ФИО , Студенты.Год_Поступления , Студенты.Форма_обучения, Специальности.Специальности ,Группы.Группа  from Студенты, Группы, Специальности where Специальности.Код_специальности =Группы.Код_специальности and Группы.Код_группы = Студенты.Код_группы and Специальности.Специальности='Программист' and Группы.Группа like '%" + s + "%' and  Студенты.Год_Поступления="+ dateTimePicker1.Value.Date.Year +"";

                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);


                dataGridView1.DataSource = dt;
                dataGridView1.Columns[0].Width = 200;
                dataGridView1.Columns[1].Width = 100;
                con.Close();
               
            }


            if (radioButton2.Checked && !(comboBox1.Text == ""))
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                string s = comboBox1.Text;

                if ("Все" == comboBox1.Text)
                {
                    s = "Б";
                }

                cmd.CommandText = "Select       Студенты.ФИО , Студенты.Год_Поступления , Студенты.Форма_обучения, Специальности.Специальности  from Студенты, Группы, Специальности where Специальности.Код_специальности =Группы.Код_специальности and Группы.Код_группы = Студенты.Код_группы and Специальности.Специальности='Бухгалтера' and Группы.Группа like '%" + s + "%' and  Студенты.Год_Поступления=" + dateTimePicker1.Value.Date.Year + " ";

                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);


                dataGridView1.DataSource = dt;
                dataGridView1.Columns[0].Width = 200;
                dataGridView1.Columns[1].Width = 100;
                con.Close();
              
            }

            if (radioButton3.Checked && !(comboBox1.Text == ""))
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                string s = comboBox1.Text;

                if ("Все" == comboBox1.Text)
                {
                    s = "И";
                }

                cmd.CommandText = "Select        Студенты.ФИО , Студенты.Год_Поступления , Студенты.Форма_обучения, Специальности.Специальности  from Студенты, Группы, Специальности where Специальности.Код_специальности =Группы.Код_специальности and Группы.Код_группы = Студенты.Код_группы and Специальности.Специальности='Экономист' and Группы.Группа like '%" + s + "%' and  Студенты.Год_Поступления=" + dateTimePicker1.Value.Date.Year + " ";

                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);


                dataGridView1.DataSource = dt;
                dataGridView1.Columns[0].Width = 200;
                dataGridView1.Columns[1].Width = 100;
                con.Close();
              
            }

            if (radioButton4.Checked && !(comboBox1.Text == ""))
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                string s = comboBox1.Text;

                if ("Все" == comboBox1.Text)
                {
                    s = "П";
                }

                cmd.CommandText = "Select       Студенты.ФИО , Студенты.Год_Поступления , Студенты.Форма_обучения, Специальности.Специальности  from Студенты, Группы, Специальности where Специальности.Код_специальности =Группы.Код_специальности and Группы.Код_группы = Студенты.Код_группы and Специальности.Специальности='Правоведы' and Группы.Группа like '%" + s + "%' and  Студенты.Год_Поступления=" + dateTimePicker1.Value.Date.Year + " ";

                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);


                dataGridView1.DataSource = dt;
                dataGridView1.Columns[0].Width = 200;
                dataGridView1.Columns[1].Width = 100;
                con.Close();
               
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            comboBox1.Items.Clear(); comboBox1.Items.AddRange(new object[] { "Все", "СП 105", "СП 205", "СП 305", "СП 405" });
        }
        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            comboBox1.Items.Clear(); comboBox1.Items.AddRange(new object[] { "Все", "П 106", "П 206", "П 306" });
        }

        private void radioButton2_CheckedChanged_1(object sender, EventArgs e)
        {

            comboBox1.Items.Clear(); comboBox1.Items.AddRange(new object[] { "Все", "Б 101", "Б 201", "Б 301" });
        }

        private void radioButton3_CheckedChanged_1(object sender, EventArgs e)
        {
            comboBox1.Items.Clear(); comboBox1.Items.AddRange(new object[] { "Все", "И 104", "И 204", "И 304" });
        }

        private void radioButton4_CheckedChanged_1(object sender, EventArgs e)
        {
            comboBox1.Items.Clear(); comboBox1.Items.AddRange(new object[] { "Все", "П 106", "П 206", "П 306" });
        }

        private void Form4_Load(object sender, EventArgs e)
        {

            dataGridView1.AllowUserToResizeColumns = false;
            dataGridView1.AllowUserToResizeRows = false;

            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "yyyy";
            dateTimePicker1.ShowUpDown = true;

            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.CustomFormat = "yyyy";
            dateTimePicker2.ShowUpDown = true;

            dateTimePicker3.Format = DateTimePickerFormat.Custom;
            dateTimePicker3.CustomFormat = "yyyy";
            dateTimePicker3.ShowUpDown = true;

            comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
           
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (radioButton8.Checked)
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;          
                cmd.CommandText = "Select   Дисциплины , Семестр, Отв_Часы,Форма_обучения from Дисциплины,  План where  Дисциплины.Код_Специальности=1 and Дисциплины.Код_Дисциплины=План.Дисциплина";
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                dataGridView1.Columns[0].Width = 300;
                dataGridView1.Columns[1].Width = 100;
                con.Close();        
            }


            if (radioButton7.Checked )
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select   Дисциплины , Семестр, Отв_Часы,Форма_обучения from Дисциплины,  План where  Дисциплины.Код_Специальности=2 and Дисциплины.Код_Дисциплины=План.Дисциплина";
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                dataGridView1.Columns[0].Width = 300;
                dataGridView1.Columns[1].Width = 100;
                con.Close();              
            }

            if (radioButton6.Checked )
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select   Дисциплины , Семестр, Отв_Часы,Форма_обучения from Дисциплины,  План where  Дисциплины.Код_Специальности=3 and Дисциплины.Код_Дисциплины=План.Дисциплина";
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                dataGridView1.Columns[0].Width = 300;
                dataGridView1.Columns[1].Width = 100;
                con.Close();             
            }

            if (radioButton5.Checked)
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select   Дисциплины , Семестр, Отв_Часы,Форма_обучения from Дисциплины,  План where  Дисциплины.Код_Специальности=4 and Дисциплины.Код_Дисциплины=План.Дисциплина";
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                dataGridView1.Columns[0].Width = 300;
                dataGridView1.Columns[1].Width = 100;
                con.Close();              
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT  Студенты.ФИО, Дисциплины.Дисциплины , Отметки.Год , Отметки.Семестр, Отметки.Оценка  from Студенты, Дисциплины , План, Отметки where Студенты.Код_Студента = Отметки.Код_Студента and Отметки.Оценка <4 and Отметки.Год =" + dateTimePicker2.Value.Date.Year + " and Дисциплины.Код_Дисциплины = План.Дисциплина and План.Код_Плана = Отметки.Код_Плана";
           
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);


            dataGridView1.DataSource = dt;
            dataGridView1.Columns[0].Width = 220;

           
            con.Close();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT  Студенты.ФИО, Дисциплины.Дисциплины , Отметки.Год , Отметки.Семестр, Отметки.Оценка  from Студенты, Дисциплины , План, Отметки where Студенты.Код_Студента = Отметки.Код_Студента and Отметки.Оценка > 7 and Отметки.Год ="+ dateTimePicker3.Value.Date.Year+" and Дисциплины.Код_Дисциплины = План.Дисциплина and План.Код_Плана = Отметки.Код_Плана";
          
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);


            dataGridView1.DataSource = dt;
            dataGridView1.Columns[0].Width = 220;

            con.Close();
        }
    }
}
