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
    public partial class Form3 : Form
    {
        OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Control_za_studentami.mdb");

        public Form3()
        {
           
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form2 qw = new Form2();
            qw.Show();
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {

           if (radioButton1.Checked)
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT COUNT(Код_Студента) FROM Студенты; ";

                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);


                dataGridView1.DataSource = dt;
                dataGridView1.Columns[0].Width = 50;
                dataGridView1.Columns[0].HeaderCell.Value = "Ответ";
               
                con.Close();
            }
            if (radioButton2.Checked)
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT Count(Код_Студента) FROM Студенты where Студенты.Код_группы = 1 or Студенты.Код_группы = 7 or Студенты.Код_группы = 13 or Студенты.Код_группы = 18 ";
              
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);


                dataGridView1.DataSource = dt;
                dataGridView1.Columns[0].Width = 50;
                dataGridView1.Columns[0].HeaderCell.Value = "Ответ";
                
                con.Close();
            }
            if (radioButton3.Checked)
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT Count(Код_Студента) FROM Студенты where Студенты.Код_группы = 4 or Студенты.Код_группы = 10 or Студенты.Код_группы = 16 ";
              
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);


                dataGridView1.DataSource = dt;
                dataGridView1.Columns[0].Width = 50;
                dataGridView1.Columns[0].HeaderCell.Value = "Ответ";
            
                con.Close();
            }
            if (radioButton4.Checked)
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT Count(Код_Студента) FROM Студенты where Студенты.Код_группы = 3 or Студенты.Код_группы = 9 or Студенты.Код_группы = 15 ";
            
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);


                dataGridView1.DataSource = dt;
                dataGridView1.Columns[0].Width = 50;
                dataGridView1.Columns[0].HeaderCell.Value = "Ответ";
               
                con.Close();
            }
            if (radioButton5.Checked)
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT Count(Код_Студента) FROM Студенты where Студенты.Код_группы = 2 or Студенты.Код_группы = 8 or Студенты.Код_группы = 14 ";
               
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);


                dataGridView1.DataSource = dt;
                dataGridView1.Columns[0].Width = 50;
                dataGridView1.Columns[0].HeaderCell.Value = "Ответ";
              
                con.Close();
            }
          
        }

        

        private void comboBox2_TextChanged(object sender, EventArgs e)

        {
            if (comboBox2.Text == "Программист") { comboBox1.Items.Clear(); comboBox1.Items.AddRange(new object[] { "СП 105", "СП 205", "СП 305", "СП 405" }); };
            if (comboBox2.Text == "Эконамист") { comboBox1.Items.Clear(); comboBox1.Items.AddRange(new object[] { "И 104", "И 204", "И 304" }); };
            if (comboBox2.Text == "Бухгалтер") { comboBox1.Items.Clear(); comboBox1.Items.AddRange(new object[] { "Б 101", "Б 201", "Б 301" }); };
            if (comboBox2.Text == "Правовед") { comboBox1.Items.Clear(); comboBox1.Items.AddRange(new object[] { "П 106", "П 206", "П 306" }); };
        }

        private void Form3_Load_1(object sender, EventArgs e)
        {
            dataGridView1.ScrollBars = ScrollBars.None;
            comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
        }

        private void button1_Click_1(object sender, EventArgs e)
        { 
        if (radioButton1.Checked)
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
        cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT COUNT(Код_Студента) FROM Студенты; ";

                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
        OleDbDataAdapter da = new OleDbDataAdapter(cmd);
        da.Fill(dt);


                dataGridView1.DataSource = dt;
                dataGridView1.Columns[0].Width = 50;
                dataGridView1.Columns[0].HeaderCell.Value = "Ответ";
              
                con.Close();
            }
            if (radioButton2.Checked)
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
    cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT Count(Код_Студента) FROM Студенты where Студенты.Код_группы = 1 or Студенты.Код_группы = 7 or Студенты.Код_группы = 13 or Студенты.Код_группы = 18 ";
           
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
    da.Fill(dt);


                dataGridView1.DataSource = dt;
                dataGridView1.Columns[0].Width = 50;
                dataGridView1.Columns[0].HeaderCell.Value = "Ответ";
               
                con.Close();
            }
            if (radioButton3.Checked)
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT Count(Код_Студента) FROM Студенты where Студенты.Код_группы = 4 or Студенты.Код_группы = 10 or Студенты.Код_группы = 16 ";
             
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
OleDbDataAdapter da = new OleDbDataAdapter(cmd);
da.Fill(dt);


                dataGridView1.DataSource = dt;
                dataGridView1.Columns[0].Width = 50;
                dataGridView1.Columns[0].HeaderCell.Value = "Ответ";
               
                con.Close();
            }
            if (radioButton4.Checked)
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT Count(Код_Студента) FROM Студенты where Студенты.Код_группы = 3 or Студенты.Код_группы = 9 or Студенты.Код_группы = 15 ";
               
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
OleDbDataAdapter da = new OleDbDataAdapter(cmd);
da.Fill(dt);


                dataGridView1.DataSource = dt;
                dataGridView1.Columns[0].Width = 50;
                dataGridView1.Columns[0].HeaderCell.Value = "Ответ";
             
                con.Close();
            }
            if (radioButton5.Checked)
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT Count(Код_Студента) FROM Студенты where Студенты.Код_группы = 2 or Студенты.Код_группы = 8 or Студенты.Код_группы = 14 ";
               
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
OleDbDataAdapter da = new OleDbDataAdapter(cmd);
da.Fill(dt);


                dataGridView1.DataSource = dt;
                dataGridView1.Columns[0].Width = 50;
                dataGridView1.Columns[0].HeaderCell.Value = "Ответ";
             
                con.Close();
            }



            

        }

        private void button4_Click(object sender, EventArgs e)
        {
            
            string g1 = comboBox1.Text;
            int g3 = 0;
        
            if (!(g1 == ""))

            {
                if (g1 == "СП 105") { g3 = 1; }
                if (g1 == "СП 205") { g3 = 7; }
                if (g1 == "СП 305") { g3 = 13; }
                if (g1 == "СП 405") { g3 = 18; }
                if (g1 == "И 104") { g3 = 3; }
                if (g1 == "И 204") { g3 = 9; }
                if (g1 == "И 304") { g3 = 15; }
                if (g1 == "Б 101") { g3 = 4; }
                if (g1 == "Б 201") { g3 = 10; }
                if (g1 == "Б 301") { g3 = 16; }
                if (g1 == "П 106") { g3 = 2; }
                if (g1 == "П 206") { g3 = 8; }
                if (g1 == "П 306") { g3 = 14; }
                try
                {

                    con.Open();
                    OleDbCommand cmd = con.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "SELECT Count(Код_Студента) FROM Студенты where Студенты.Код_группы ="+g3+" ";
                  
                    cmd.ExecuteNonQuery();
                    DataTable dt = new DataTable();
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(dt);


                    dataGridView1.DataSource = dt;
                    dataGridView1.Columns[0].Width = 50;
                    dataGridView1.Columns[0].HeaderCell.Value = "Ответ";
                  
                    con.Close();

                }
                catch { MessageBox.Show("Неизвесная ошибка"); }

            }
            else
            {
              
                MessageBox.Show("Вы не коректно ввели данные");
            }
        }
    }
}
