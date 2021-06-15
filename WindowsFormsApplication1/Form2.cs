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
    public partial class Form2 : Form
    {
        OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Control_za_studentami.mdb");
        public Form2()
        {
            InitializeComponent();
           
            button3.FlatAppearance.BorderSize = 0;
            button3.FlatStyle = FlatStyle.Flat;
            textBox2.Text = "Поиск...";
        }
        public string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Control_za_studentami.mdb";

       

        private void button3_Click(object sender, EventArgs e)
        {
            string s = textBox1.Text;
            string g = comboBox2.Text;
            string g1 = comboBox5.Text;
            int g3 = 0;            
            if (!(s.Count() > 35)&&(s != "")&&!(g1==""))      
            {
                try {                  
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
                  
                    StreamReader f = new StreamReader("text1.txt");
                    int ss = Convert.ToInt32(f.ReadToEnd());
                    f.Close();
                    StreamWriter ff = new StreamWriter("text1.txt");

                    ss++;
                    ff.WriteLine(ss);
                    ff.Close();

                    con.Open();
                    OleDbCommand cmd = con.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "insert into Студенты values (" + ss + ",'" + s + "'," + dateTimePicker1.Value.Date.Year + ",'" + g + "'," + g3 + ")";
                    cmd.ExecuteNonQuery();
                    con.Close();
                    MessageBox.Show("Пользователь " + s + " благополучно добавлен!!");
                }
                catch
                {
                    MessageBox.Show("Неизвесная ошибка");
                }
               
            }
            else
            {
                textBox1.Clear();
                MessageBox.Show("Вы не коректно ввели данные");
            }
  
        }
       
        private void Form2_Load(object sender, EventArgs e)
        {

            dataGridView1.AllowUserToResizeColumns = false;
            dataGridView1.AllowUserToResizeRows = false;

            dataGridView2.AllowUserToResizeColumns = false;
            dataGridView2.AllowUserToResizeRows = false;

            dataGridView1.ScrollBars = ScrollBars.Vertical;
            dataGridView2.ScrollBars = ScrollBars.Vertical;
            comboBox3.SelectedItem = "дневная";
            comboBox4.SelectedItem = "Программист";
            comboBox5.SelectedItem = "СП 105";
            comboBox2.SelectedItem = "дневная";
            comboBox1.SelectedItem = "Программист";
            comboBox7.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox8.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox9.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox4.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox5.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox6.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox13.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox11.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox10.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBox12.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "yyyy";
            dateTimePicker1.ShowUpDown = true;
            dateTimePicker4.Format = DateTimePickerFormat.Custom;
            dateTimePicker4.CustomFormat = "yyyy";
            dateTimePicker4.ShowUpDown = true;
            dateTimePicker3.Format = DateTimePickerFormat.Custom;
            dateTimePicker3.CustomFormat = "yyyy";
            dateTimePicker3.ShowUpDown = true;
            dateTimePicker5.Format = DateTimePickerFormat.Custom;
            dateTimePicker5.CustomFormat = "yyyy";
            dateTimePicker5.ShowUpDown = true;
            con.Open();


            dateTimePicker6.Format = DateTimePickerFormat.Custom;
            dateTimePicker6.CustomFormat = "yyyy";
            dateTimePicker6.ShowUpDown = true;
            
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "Select  Код_Студента ,      Студенты.ФИО , Студенты.Год_Поступления , Студенты.Форма_обучения, Специальности.Специальности  from Студенты, Группы, Специальности where Специальности.Код_специальности =Группы.Код_специальности and Группы.Код_группы = Студенты.Код_группы";
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);
            dataGridView1.DataSource = dt;
            dataGridView1.Columns[0].Width = 200;
            con.Close();
           
            string queryString = "SELECT Sum(Оценка) FROM Отметки where год=2019";
            string queryString1 = "SELECT count(Код_успеваемости_студента) FROM Отметки where год=2019";

            string queryString2 = "SELECT Sum(Оценка) FROM Отметки where год=2018";
            string queryString22 = "SELECT count(Код_успеваемости_студента) FROM Отметки where год=2018";

            string queryString3 = "SELECT Sum(Оценка) FROM Отметки where год=2017";
            string queryString33 = "SELECT count(Код_успеваемости_студента) FROM Отметки where год=2017";
            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    OleDbCommand command = new OleDbCommand(queryString, connection);
                    OleDbCommand command1 = new OleDbCommand(queryString1, connection);

                    OleDbCommand command2 = new OleDbCommand(queryString2, connection);
                    OleDbCommand command22 = new OleDbCommand(queryString22, connection);

                    OleDbCommand command3 = new OleDbCommand(queryString3, connection);
                    OleDbCommand command33 = new OleDbCommand(queryString33, connection);
                    connection.Open();
                    ////2019
                    String ss = command.ExecuteScalar().ToString();
                    String ss1 = command1.ExecuteScalar().ToString();
                    int a = Convert.ToInt32(ss);
                    int aa = Convert.ToInt32(ss1);
                    int q = 0;
                    q = a / aa;
                    label27.Text = Convert.ToString(q);
                    ////2018
                    String ss2 = command2.ExecuteScalar().ToString();
                    String ss22 = command22.ExecuteScalar().ToString();
                    int a2 = Convert.ToInt32(ss2);
                    int aa2 = Convert.ToInt32(ss22);
                    int q2 = 0;
                    q2 = a2 / aa2;
                    label33.Text = Convert.ToString(q2);
                    ////2017
                    String ss3 = command3.ExecuteScalar().ToString();
                    String ss33 = command33.ExecuteScalar().ToString();
                    int a3 = Convert.ToInt32(ss3);
                    int aa3 = Convert.ToInt32(ss33);
                    int q3 = 0;
                    q3 = a3 / aa3;
                    label34.Text = Convert.ToString(q3);
                }
            }


            catch (Exception ex)
            {
                MessageBox.Show(" 1 failed to connect");
            }
           
        }
        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
          
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "DELETE FROM Студенты WHERE Код_Студента =" + (Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value)) + "";

            cmd.ExecuteNonQuery();
           
            con.Close();
            }
            catch { }
        }

    
        private void textBox2_TextChanged_1(object sender, EventArgs e)
        {
            try
            {

                int q = 1;
            string ee = textBox2.Text;
            do
            {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select Студенты.ФИО , Студенты.Год_Поступления , Студенты.Форма_обучения, Специальности.Специальности  from Студенты, Группы, Специальности where Специальности.Код_специальности =Группы.Код_специальности and Группы.Код_группы = Студенты.Код_группы and ФИО like '%" + ee + "%'";

                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);


                dataGridView1.DataSource = dt;
                
                dataGridView1.Columns[0].Width = 200;

               
                con.Close();

            }
            while (q == 2);
            }
            catch { }
        }

     
        private void button7_Click(object sender, EventArgs e)

        {
            try { 
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "Select  Код_Студента ,Студенты.ФИО , Студенты.Год_Поступления , Студенты.Форма_обучения, Специальности.Специальности  from Студенты, Группы, Специальности where Специальности.Код_специальности =Группы.Код_специальности and Группы.Код_группы = Студенты.Код_группы";

            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);


            dataGridView1.DataSource = dt;
            dataGridView1.Columns[0].Width = 200;

            // dataGridView1["Column1", 1].Value = "ddddd";
            con.Close();
        }
            catch { }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            Form3 qw = new Form3();
            qw.Show();
            this.Close();

        
                 }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                textBox3.Text = " " + (Convert.ToString(dataGridView1.CurrentRow.Cells[1].Value)) + "  ";
            }
            catch { }

        }

        private void button13_Click(object sender, EventArgs e)
        {
            string s = textBox3.Text;
            string g = comboBox3.Text;
            string g1 = comboBox6.Text;
            int g3 = 0;
           
           if (!(s.Count() > 35) && (s != "")&&!(g1==""))

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
                try {

                    con.Open();
                    OleDbCommand cmd = con.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = " UPDATE Студенты SET ФИО='" + s + "' , Год_Поступления=" + dateTimePicker2.Value.Date.Year + ", Форма_обучения='" + g + "', Код_группы=" + g3 + "  WHERE Код_Студента =" + (Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value)) + " ";
                    //      UPDATE users SET role = 'модератор' WHERE id_user = 1
                    //      insert into Студенты values (" + ss + ",'" + s + "'," + dateTimePicker1.Value.Date.Year + ",'" + g + "'," + g3 + ")
                    cmd.ExecuteNonQuery();
                    con.Close();
                    MessageBox.Show("Пользователь " + s + " благополучно отредактирован!!");
                }
                catch { MessageBox.Show("Неизвесная ошибка"); }

            }
            else
            {
                textBox1.Clear();
               MessageBox.Show("Вы не коректно ввели данные");
          }
        }

    
        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text == "Программист") { comboBox5.Items.Clear(); comboBox5.Items.AddRange(new object[] { "СП 105", "СП 205", "СП 305", "СП 405" }); };
            if (comboBox1.Text == "Эконамист") { comboBox5.Items.Clear(); comboBox5.Items.AddRange(new object[] { "И 104", "И 204", "И 304" }); };
            if (comboBox1.Text == "Бухгалтер") { comboBox5.Items.Clear(); comboBox5.Items.AddRange(new object[] { "Б 101", "Б 201", "Б 301" }); };
            if (comboBox1.Text == "Правовед") { comboBox5.Items.Clear(); comboBox5.Items.AddRange(new object[] { "П 106", "П 206", "П 306" }); };
        
        }

        private void comboBox4_TextChanged(object sender, EventArgs e)
        {
            if (comboBox4.Text == "Программист") { comboBox6.Items.Clear(); comboBox6.Items.AddRange(new object[] { "СП 105", "СП 205", "СП 305", "СП 405" }); };
            if (comboBox4.Text == "Эконамист") { comboBox6.Items.Clear(); comboBox6.Items.AddRange(new object[] { "И 104", "И 204", "И 304" }); };
            if (comboBox4.Text == "Бухгалтер") { comboBox6.Items.Clear(); comboBox6.Items.AddRange(new object[] { "Б 101", "Б 201", "Б 301" }); };
            if (comboBox4.Text == "Правовед") { comboBox6.Items.Clear(); comboBox6.Items.AddRange(new object[] { "П 106", "П 206", "П 306" }); };
        }

 
        private void button6_Click_1(object sender, EventArgs e)
        {
         
            string queryString3 = "SELECT Sum(Оценка) FROM Отметки where год=" + dateTimePicker5.Value.Date.Year + "";
            string queryString33 = "SELECT count(Код_успеваемости_студента) FROM Отметки where год=" + dateTimePicker5.Value.Date.Year + "";
            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    OleDbCommand command = new OleDbCommand(queryString3, connection);
                    OleDbCommand command1 = new OleDbCommand(queryString33, connection);

                   
                    connection.Open();
                    ////2019
                    String ss = command.ExecuteScalar().ToString();
                    String ss1 = command1.ExecuteScalar().ToString();
                    int a = Convert.ToInt32(ss);
                    int aa = Convert.ToInt32(ss1);
                    int q = 0;
                    q = a / aa;
                    label35.Text = Convert.ToString(q);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("failed to connect");
            }

        }


        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try { 
            con.Open();
            OleDbCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT  Дисциплины.Дисциплины , Отметки.Год , Отметки.Семестр, Отметки.Оценка ,Отметки.Код_успеваемости_студента from Дисциплины , План, Отметки where Отметки.Код_Студента=" + (Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value)) + " and Дисциплины.Код_Дисциплины=План.Дисциплина and План.Код_Плана=Отметки.Код_Плана";

            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);


            dataGridView2.DataSource = dt;
            dataGridView2.Columns[0].Width = 220;

         
            con.Close();

       
        }
            catch { }
            comboBox8.Items.Clear();
            comboBox7.Items.Clear();
            try {
                string queryString5 = "SELECT Sum(Оценка) FROM Отметки where год=2019 and Отметки.Код_студента=" + (Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value)) + "";

                string queryString55 = "SELECT count(Код_успеваемости_студента) FROM Отметки where год=2019 and Код_студента=" + (Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value)) + " ";
            
                try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    OleDbCommand command55 = new OleDbCommand(queryString5, connection);
                    OleDbCommand command555 = new OleDbCommand(queryString55, connection);


                    connection.Open();
                    ////2019
                    String ss = command55.ExecuteScalar().ToString();
                    String ss1 = command555.ExecuteScalar().ToString();
                    int a = Convert.ToInt32(ss);
                    int aa = Convert.ToInt32(ss1);
                    double q = 0;
                    q = a / aa;
                    label39.Text = Convert.ToString(q);


                }


            }
            catch (Exception ex)
            {
                label39.Text = "null";
                //MessageBox.Show("failed to connect");
            }
            }
            catch { }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            string s = comboBox9.Text;
            string g = comboBox8.Text;
            string g1 = comboBox7.Text;

            int g3 = 0;



            if ((g != "") && !(g1 == "") && !(s == ""))

            {
                try
                {

                    if (g1 == "Матемотическое моделирование(дневная)") { g3 = 1; }
                    if (g1 == "Защита компьютерной информации (дневная)") { g3 = 3; }
                    if (g1 == "Базы данных и системы управления базами данных (дневная)") { g3 = 5; }
                    if (g1 == "Конструирование программ и языки программирования (дневная)") { g3 = 7; }
                    if (g1 == "Системное программирование (дневная)") { g3 = 9; }
                    if (g1 == "Системное программное обеспечение (дневная)") { g3 = 11; }
                    if (g1 == "Матемотическое моделирование(заочная)") { g3 = 2; }
                    if (g1 == "Защита компьютерной информации (заочная)") { g3 = 4; }
                    if (g1 == "Базы данных и системы управления базами данных (заочная)") { g3 = 6; }
                    if (g1 == "Конструирование программ и языки программирования (заочная)") { g3 = 8; }
                    if (g1 == "Системное программирование (заочная)") { g3 = 10; }
                    if (g1 == "Системное программное обеспечение (заочная)") { g3 = 12; }


                    if (g1 == "Охрана труда (дневная)") { g3 = 25; }
                    if (g1 == "Основы охраны труда (дневная)") { g3 = 27; }
                    if (g1 == "Теория бухгалтерского учета (дневная)") { g3 = 29; }
                    if (g1 == "Ревизия и контроль (дневная)") { g3 = 31; }
                    if (g1 == "Логистика (дневная)") { g3 = 33; }
                    if (g1 == "Основы маркетинга (дневная)") { g3 = 35; }
                    if (g1 == "Охрана труда (заочная)") { g3 = 26; }
                    if (g1 == "Основы охраны труда (заочная)") { g3 = 28; }
                    if (g1 == "Теория бухгалтерского учета (заочная)") { g3 = 30; }
                    if (g1 == "Ревизия и контроль (заочная)") { g3 = 32; }
                    if (g1 == "Логистика (заочная)") { g3 = 34; }
                    if (g1 == "Основы маркетинга (заочная)") { g3 = 36; }

                    if (g1 == "Статистика (дневная)") { g3 = 37; }
                    if (g1 == "Налогообложение (дневная)") { g3 = 39; }
                    if (g1 == "Основы экономики (дневная)") { g3 = 41; }
                    if (g1 == "Экономика Торговли (дневная)") { g3 = 43; }
                    if (g1 == "Статистика (заочная)") { g3 = 38; }
                    if (g1 == "Налогообложение (заочная)") { g3 = 41; }
                    if (g1 == "Основы экономики (заочная)") { g3 = 42; }
                    if (g1 == "Экономика Торговли (заочная)") { g3 = 44; }


                    if (g1 == "Административное право (дневная)") { g3 = 15; }
                    if (g1 == "Логика (дневная)") { g3 = 17; }
                    if (g1 == "Земельное право (дневная)") { g3 = 19; }
                    if (g1 == "Жилищное право (дневная)") { g3 = 21; }
                    if (g1 == "Налоговое право (дневная)") { g3 = 23; }
                    if (g1 == "Таможенное право (дневная)") { g3 = 25; }
                    if (g1 == "Административное право (заочная)") { g3 = 15; }
                    if (g1 == "Логика (заочная)") { g3 = 17; }
                    if (g1 == "Земельное право (заочная)") { g3 = 19; }
                    if (g1 == "Жилищное право (заочная)") { g3 = 21; }
                    if (g1 == "Налоговое право (заочная)") { g3 = 23; }
                    if (g1 == "Таможенное право (заочная)") { g3 = 25; }



                    StreamReader f = new StreamReader("text2.txt");
                    int ss = Convert.ToInt32(f.ReadToEnd());
                    f.Close();
                    StreamWriter ff = new StreamWriter("text2.txt");

                    ss++;
                    ff.WriteLine(ss);
                    ff.Close();


                    con.Open();
                    OleDbCommand cmd = con.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "insert into Отметки values (" + ss + "," + (Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value)) + "," + dateTimePicker3.Value.Date.Year + "," + (Convert.ToInt32(comboBox8.Text)) + "," + (Convert.ToInt32(comboBox9.Text)) + "," + g3 + ")";
                    cmd.ExecuteNonQuery();
                    con.Close();
                    MessageBox.Show("Отметка благополучно добавлена!!");
                }
                catch { }
             } 

            
            else
            {
                textBox1.Clear();
                MessageBox.Show("Вы не коректно ввели данные");
            }
                

        }

        private void button16_Click(object sender, EventArgs e)

        {
            try {
                String ss = (Convert.ToString(dataGridView1.CurrentRow.Cells[4].Value));
                String ss1 = (Convert.ToString(dataGridView1.CurrentRow.Cells[3].Value));

                comboBox8.Items.Clear();
                comboBox7.Items.Clear();
                if ((ss == "Программист") && (ss1 == "дневная"))
                {
                    comboBox8.Items.Clear(); comboBox8.Items.AddRange(new object[] { "1", "2", "3", "4", "5", "6", "7", "8" });
                    comboBox7.Items.Clear(); comboBox7.Items.AddRange(new object[] { "Матемотическое моделирование (дневная)", "Защита компьютерной информации (дневная)", "Базы данных и системы управления базами данных (дневная)", "Конструирование программ и языки программирования (дневная)", "Системное программирование (дневная)", "Системное программное обеспечение (дневная)" });
                }
                if ((ss == "Программист") && (ss1 == "заочная"))
                {
                    comboBox8.Items.Clear(); comboBox8.Items.AddRange(new object[] { "1", "2", "3", "4", "5", "6", "7", "8" });
                    comboBox7.Items.Clear(); comboBox7.Items.AddRange(new object[] { "Матемотическое моделирование (заочная)", "Защита компьютерной информации (заочная)", "Базы данных и системы управления базами данных (заочная)", "Конструирование программ и языки программирования (заочная)", "Системное программирование (заочная)", "Системное программное обеспечение (заочная)" });
                }


                if ((ss == "Экономист") && (ss1 == "дневная"))
                {
                    comboBox8.Items.Clear(); comboBox8.Items.AddRange(new object[] { "1", "2", "3", "4", "5", "6" });
                    comboBox7.Items.Clear(); comboBox7.Items.AddRange(new object[] { "Охрана труда (дневная)", "Основы охраны труда (дневная)", "Теория бухгалтерского учета (дневная)", "Ревизия и контроль (дневная)", "Логистика (дневная)", "Основы маркетинга (дневная)" });
                }
                if ((ss == "Экономист") && (ss1 == "заочная"))
                {
                    comboBox8.Items.Clear(); comboBox8.Items.AddRange(new object[] { "1", "2", "3", "4", "5", "6" });
                    comboBox7.Items.Clear(); comboBox7.Items.AddRange(new object[] { "Охрана труда (заочная)", "Основы охраны труда (заочная)", "Теория бухгалтерского учета (заочная)", "Ревизия и контроль (заочная)", "Логистика (заочная)", "Основы маркетинга (заочная)" });
                }


                if ((ss == "Бухгалтера") && (ss1 == "дневная"))
                {
                    comboBox8.Items.Clear(); comboBox8.Items.AddRange(new object[] { "1", "2", "3", "4", "5", "6" });
                    comboBox7.Items.Clear(); comboBox7.Items.AddRange(new object[] { "Статистика (дневная)", "Налогообложение (дневная)", "Основы экономики (дневная)", "Экономика Торговли (дневная)" });
                }
                if ((ss == "Бухгалтера") && (ss1 == "заочная"))
                {
                    comboBox8.Items.Clear(); comboBox8.Items.AddRange(new object[] { "1", "2", "3", "4", "5", "6" });
                    comboBox7.Items.Clear(); comboBox7.Items.AddRange(new object[] { "Статистика (заочная)", "Налогообложение (заочная)", "Основы экономики (заочная)", "Экономика Торговли (заочная)" });
                }



                if ((ss == "Правоведы") && (ss1 == "дневная"))
                {
                    comboBox8.Items.Clear(); comboBox8.Items.AddRange(new object[] { "1", "2", "3", "4", "5", "6" });
                    comboBox7.Items.Clear(); comboBox7.Items.AddRange(new object[] { "Административное право (дневная)", "Логика (дневная)", "Земельное право (дневная)", "Жилищное право (дневная)", "Налоговое право (дневная)", "Таможенное право (дневная)" });
                }
                if ((ss == "Правоведы") && (ss1 == "заочная"))
                {
                    comboBox8.Items.Clear(); comboBox8.Items.AddRange(new object[] { "1", "2", "3", "4", "5", "6" });
                    comboBox7.Items.Clear(); comboBox7.Items.AddRange(new object[] { "Административное право (заочная)", "Логика (заочная)", "Земельное право (заочная)", "Жилищное право (заочная)", "Налоговое право (заочная)", "Таможенное право (заочная)" });
                }
            }
            catch { }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            try {
                String ss = (Convert.ToString(dataGridView1.CurrentRow.Cells[4].Value));
                String ss1 = (Convert.ToString(dataGridView1.CurrentRow.Cells[3].Value));

                comboBox11.Items.Clear();
                comboBox12.Items.Clear();
                if ((ss == "Программист") && (ss1 == "дневная"))
                {
                    comboBox11.Items.Clear(); comboBox11.Items.AddRange(new object[] { "1", "2", "3", "4", "5", "6", "7", "8" });
                    comboBox12.Items.Clear(); comboBox12.Items.AddRange(new object[] { "Матемотическое моделирование (дневная)", "Защита компьютерной информации (дневная)", "Базы данных и системы управления базами данных (дневная)", "Конструирование программ и языки программирования (дневная)", "Системное программирование (дневная)", "Системное программное обеспечение (дневная)" });
                }
                if ((ss == "Программист") && (ss1 == "заочная"))
                {
                    comboBox11.Items.Clear(); comboBox11.Items.AddRange(new object[] { "1", "2", "3", "4", "5", "6", "7", "8" });
                    comboBox12.Items.Clear(); comboBox12.Items.AddRange(new object[] { "Матемотическое моделирование (заочная)", "Защита компьютерной информации (заочная)", "Базы данных и системы управления базами данных (заочная)", "Конструирование программ и языки программирования (заочная)", "Системное программирование (заочная)", "Системное программное обеспечение (заочная)" });
                }


                if ((ss == "Экономист") && (ss1 == "дневная"))
                {
                    comboBox11.Items.Clear(); comboBox11.Items.AddRange(new object[] { "1", "2", "3", "4", "5", "6" });
                    comboBox12.Items.Clear(); comboBox12.Items.AddRange(new object[] { "Охрана труда (дневная)", "Основы охраны труда (дневная)", "Теория бухгалтерского учета (дневная)", "Ревизия и контроль (дневная)", "Логистика (дневная)", "Основы маркетинга (дневная)" });
                }
                if ((ss == "Экономист") && (ss1 == "заочная"))
                {
                    comboBox11.Items.Clear(); comboBox11.Items.AddRange(new object[] { "1", "2", "3", "4", "5", "6" });
                    comboBox12.Items.Clear(); comboBox12.Items.AddRange(new object[] { "Охрана труда (заочная)", "Основы охраны труда (заочная)", "Теория бухгалтерского учета (заочная)", "Ревизия и контроль (заочная)", "Логистика (заочная)", "Основы маркетинга (заочная)" });
                }


                if ((ss == "Бухгалтера") && (ss1 == "дневная"))
                {
                    comboBox11.Items.Clear(); comboBox11.Items.AddRange(new object[] { "1", "2", "3", "4", "5", "6" });
                    comboBox12.Items.Clear(); comboBox12.Items.AddRange(new object[] { "Статистика (дневная)", "Налогообложение (дневная)", "Основы экономики (дневная)", "Экономика Торговли (дневная)" });
                }
                if ((ss == "Бухгалтера") && (ss1 == "заочная"))
                {
                    comboBox11.Items.Clear(); comboBox11.Items.AddRange(new object[] { "1", "2", "3", "4", "5", "6" });
                    comboBox12.Items.Clear(); comboBox12.Items.AddRange(new object[] { "Статистика (заочная)", "Налогообложение (заочная)", "Основы экономики (заочная)", "Экономика Торговли (заочная)" });
                }


                if ((ss == "Правоведы") && (ss1 == "дневная"))
                {
                    comboBox11.Items.Clear(); comboBox11.Items.AddRange(new object[] { "1", "2", "3", "4", "5", "6" });
                    comboBox12.Items.Clear(); comboBox12.Items.AddRange(new object[] { "Административное право (дневная)", "Логика (дневная)", "Земельное право (дневная)", "Жилищное право (дневная)", "Налоговое право (дневная)", "Таможенное право (дневная)" });
                }
                if ((ss == "Правоведы") && (ss1 == "заочная"))
                {
                    comboBox11.Items.Clear(); comboBox11.Items.AddRange(new object[] { "1", "2", "3", "4", "5", "6" });
                    comboBox12.Items.Clear(); comboBox12.Items.AddRange(new object[] { "Административное право (заочная)", "Логика (заочная)", "Земельное право (заочная)", "Жилищное право (заочная)", "Налоговое право (заочная)", "Таможенное право (заочная)" });
                }
            }
            catch { }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            try {
                string s = comboBox10.Text;
                string g = comboBox11.Text;
                string g1 = comboBox12.Text;

                int g3 = 0;


                if ((g != "") && !(g1 == "") && !(s == ""))

                {
            
                    if (g1 == "Матемотическое моделирование(дневная)") { g3 = 1; }
                    if (g1 == "Защита компьютерной информации (дневная)") { g3 = 3; }
                    if (g1 == "Базы данных и системы управления базами данных (дневная)") { g3 = 5; }
                    if (g1 == "Конструирование программ и языки программирования (дневная)") { g3 = 7; }
                    if (g1 == "Системное программирование (дневная)") { g3 = 9; }
                    if (g1 == "Системное программное обеспечение (дневная)") { g3 = 11; }
                    if (g1 == "Матемотическое моделирование(заочная)") { g3 = 2; }
                    if (g1 == "Защита компьютерной информации (заочная)") { g3 = 4; }
                    if (g1 == "Базы данных и системы управления базами данных (заочная)") { g3 = 6; }
                    if (g1 == "Конструирование программ и языки программирования (заочная)") { g3 = 8; }
                    if (g1 == "Системное программирование (заочная)") { g3 = 10; }
                    if (g1 == "Системное программное обеспечение (заочная)") { g3 = 12; }


                    if (g1 == "Охрана труда (дневная)") { g3 = 25; }
                    if (g1 == "Основы охраны труда (дневная)") { g3 = 27; }
                    if (g1 == "Теория бухгалтерского учета (дневная)") { g3 = 29; }
                    if (g1 == "Ревизия и контроль (дневная)") { g3 = 31; }
                    if (g1 == "Логистика (дневная)") { g3 = 33; }
                    if (g1 == "Основы маркетинга (дневная)") { g3 = 35; }
                    if (g1 == "Охрана труда (заочная)") { g3 = 26; }
                    if (g1 == "Основы охраны труда (заочная)") { g3 = 28; }
                    if (g1 == "Теория бухгалтерского учета (заочная)") { g3 = 30; }
                    if (g1 == "Ревизия и контроль (заочная)") { g3 = 32; }
                    if (g1 == "Логистика (заочная)") { g3 = 34; }
                    if (g1 == "Основы маркетинга (заочная)") { g3 = 36; }

                    if (g1 == "Статистика (дневная)") { g3 = 37; }
                    if (g1 == "Налогообложение (дневная)") { g3 = 39; }
                    if (g1 == "Основы экономики (дневная)") { g3 = 41; }
                    if (g1 == "Экономика Торговли (дневная)") { g3 = 43; }
                    if (g1 == "Статистика (заочная)") { g3 = 38; }
                    if (g1 == "Налогообложение (заочная)") { g3 = 41; }
                    if (g1 == "Основы экономики (заочная)") { g3 = 42; }
                    if (g1 == "Экономика Торговли (заочная)") { g3 = 44; }


                    if (g1 == "Административное право (дневная)") { g3 = 15; }
                    if (g1 == "Логика (дневная)") { g3 = 17; }
                    if (g1 == "Земельное право (дневная)") { g3 = 19; }
                    if (g1 == "Жилищное право (дневная)") { g3 = 21; }
                    if (g1 == "Налоговое право (дневная)") { g3 = 23; }
                    if (g1 == "Таможенное право (дневная)") { g3 = 25; }
                    if (g1 == "Административное право (заочная)") { g3 = 15; }
                    if (g1 == "Логика (заочная)") { g3 = 17; }
                    if (g1 == "Земельное право (заочная)") { g3 = 19; }
                    if (g1 == "Жилищное право (заочная)") { g3 = 21; }
                    if (g1 == "Налоговое право (заочная)") { g3 = 23; }
                    if (g1 == "Таможенное право (заочная)") { g3 = 25; }



                    StreamReader f = new StreamReader("text2.txt");
                    int ss = Convert.ToInt32(f.ReadToEnd());
                    f.Close();
                    StreamWriter ff = new StreamWriter("text2.txt");

                    ss++;
                    ff.WriteLine(ss);
                    ff.Close();


                    con.Open();
                    OleDbCommand cmd = con.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    // cmd.CommandText = "insert into Отметки values (" + ss + "," + (Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value)) + "," + dateTimePicker2.Value.Date.Year + "," + (Convert.ToInt32(comboBox8.Text)) + "," + (Convert.ToInt32(comboBox9.Text)) + "," + g3 + ")";
                    cmd.CommandText = "UPDATE Отметки SET Код_студента=" + (Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value)) + " , Год=" + dateTimePicker3.Value.Date.Year + ", Семестр=" + g + ",Оценка=" + (Convert.ToInt32(comboBox10.Text)) + ",Код_плана=" + g3 + "  WHERE Код_успеваемости_студента = " + (Convert.ToInt32(dataGridView2.CurrentRow.Cells[4].Value)) + " ";
                    cmd.ExecuteNonQuery();
                    con.Close();
                    MessageBox.Show("Отметка благополучно добавлена!!");
                }
               
                else
                {
                    textBox1.Clear();
                    MessageBox.Show("Вы не коректно ввели данные");
                }
            }
            catch { }
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            comboBox11.Items.Clear();
            comboBox12.Items.Clear();
        }

       

        private void button19_Click(object sender, EventArgs e)
        {
            try {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "DELETE FROM Отметки WHERE Код_успеваемости_студента =" + (Convert.ToInt32(dataGridView2.CurrentRow.Cells[4].Value)) + "";

                cmd.ExecuteNonQuery();

                con.Close();
            }
            catch { }
        }

        private void button22_Click(object sender, EventArgs e)
        {
            try {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT  Дисциплины.Дисциплины , Отметки.Год , Отметки.Семестр, Отметки.Оценка ,Отметки.Код_успеваемости_студента from Дисциплины , План, Отметки where Отметки.Код_Студента=" + (Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value)) + " and Дисциплины.Код_Дисциплины=План.Дисциплина and План.Код_Плана=Отметки.Код_Плана";
                
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);


                dataGridView2.DataSource = dt;
                dataGridView2.Columns[0].Width = 220;

                
                con.Close();

                comboBox8.Items.Clear();
                comboBox7.Items.Clear();
            }
            catch { }
        }

        

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            comboBox13.Items.Clear(); comboBox13.Items.AddRange(new object[] { "Все", "СП 105", "СП 205", "СП 305", "СП 405"});           
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            comboBox13.Items.Clear(); comboBox13.Items.AddRange(new object[] { "Все",  "Б 101", "Б 201", "Б 301" });
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            comboBox13.Items.Clear(); comboBox13.Items.AddRange(new object[] { "Все", "И 104", "И 204", "И 304" });
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            comboBox13.Items.Clear(); comboBox13.Items.AddRange(new object[] { "Все", "П 106", "П 206", "П 306" });
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            try {

                if (radioButton1.Checked && !(comboBox13.Text == ""))
                {
                    con.Open();
                    OleDbCommand cmd = con.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    string s = comboBox13.Text;

                    if ("Все" == comboBox13.Text)
                    {
                        s = "СП";
                    }

                    cmd.CommandText = "Select  Код_Студента ,      Студенты.ФИО , Студенты.Год_Поступления , Студенты.Форма_обучения, Специальности.Специальности  from Студенты, Группы, Специальности where Специальности.Код_специальности =Группы.Код_специальности and Группы.Код_группы = Студенты.Код_группы and Специальности.Специальности='Программист' and Группы.Группа like '%" + s + "%' ";

                    cmd.ExecuteNonQuery();
                    DataTable dt = new DataTable();
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(dt);


                    dataGridView1.DataSource = dt;
                    dataGridView1.Columns[0].Width = 250;
                    dataGridView1.Columns[1].Width = 220;
                    con.Close();
                    label23.Text = "" + comboBox13.Text + " Программисты";
                }


                if (radioButton2.Checked && !(comboBox13.Text == ""))
                {
                    con.Open();
                    OleDbCommand cmd = con.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    string s = comboBox13.Text;

                    if ("Все" == comboBox13.Text)
                    {
                        s = "Б";
                    }

                    cmd.CommandText = "Select  Код_Студента ,      Студенты.ФИО , Студенты.Год_Поступления , Студенты.Форма_обучения, Специальности.Специальности  from Студенты, Группы, Специальности where Специальности.Код_специальности =Группы.Код_специальности and Группы.Код_группы = Студенты.Код_группы and Специальности.Специальности='Бухгалтера' and Группы.Группа like '%" + s + "%' ";

                    cmd.ExecuteNonQuery();
                    DataTable dt = new DataTable();
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(dt);


                    dataGridView1.DataSource = dt;
                    dataGridView1.Columns[0].Width = 250;
                    dataGridView1.Columns[1].Width = 220;
                    con.Close();
                    label23.Text = "" + comboBox13.Text + " Бухгалтера";
                }

                if (radioButton3.Checked && !(comboBox13.Text == ""))
                {
                    con.Open();
                    OleDbCommand cmd = con.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    string s = comboBox13.Text;

                    if ("Все" == comboBox13.Text)
                    {
                        s = "И";
                    }

                    cmd.CommandText = "Select  Код_Студента ,      Студенты.ФИО , Студенты.Год_Поступления , Студенты.Форма_обучения, Специальности.Специальности  from Студенты, Группы, Специальности where Специальности.Код_специальности =Группы.Код_специальности and Группы.Код_группы = Студенты.Код_группы and Специальности.Специальности='Экономист' and Группы.Группа like '%" + s + "%' ";

                    cmd.ExecuteNonQuery();
                    DataTable dt = new DataTable();
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(dt);


                    dataGridView1.DataSource = dt;
                    dataGridView1.Columns[0].Width = 250;
                    dataGridView1.Columns[1].Width = 220;
                    con.Close();
                    label23.Text = "" + comboBox13.Text + " Экономисты";
                }

                if (radioButton4.Checked && !(comboBox13.Text == ""))
                {
                    con.Open();
                    OleDbCommand cmd = con.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    string s = comboBox13.Text;

                    if ("Все" == comboBox13.Text)
                    {
                        s = "П";
                    }

                    cmd.CommandText = "Select  Код_Студента ,      Студенты.ФИО , Студенты.Год_Поступления , Студенты.Форма_обучения, Специальности.Специальности  from Студенты, Группы, Специальности where Специальности.Код_специальности =Группы.Код_специальности and Группы.Код_группы = Студенты.Код_группы and Специальности.Специальности='Правоведы' and Группы.Группа like '%" + s + "%' ";

                    cmd.ExecuteNonQuery();
                    DataTable dt = new DataTable();
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(dt);


                    dataGridView1.DataSource = dt;
                    dataGridView1.Columns[0].Width = 250;
                    dataGridView1.Columns[1].Width = 220;
                    con.Close();
                    label23.Text = "" + comboBox13.Text + " Правоведы";
                }
            }
            catch { }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox2.Text = "";
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            try {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select  Код_Студента ,      Студенты.ФИО , Студенты.Год_Поступления , Студенты.Форма_обучения, Специальности.Специальности  from Студенты, Группы, Специальности where Специальности.Код_специальности =Группы.Код_специальности and Группы.Код_группы = Студенты.Код_группы ORDER BY Студенты.ФИО  ";

                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);


                dataGridView1.DataSource = dt;
                dataGridView1.Columns[0].Width = 250;
                dataGridView1.Columns[1].Width = 220;
                con.Close();
                label23.Text = "ФИО по возрастанию";
            }
            catch { }
        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            try {

                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select  Код_Студента ,      Студенты.ФИО , Студенты.Год_Поступления , Студенты.Форма_обучения, Специальности.Специальности  from Студенты, Группы, Специальности where Специальности.Код_специальности =Группы.Код_специальности and Группы.Код_группы = Студенты.Код_группы ORDER BY Студенты.ФИО DESC ";

                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);


                dataGridView1.DataSource = dt;
                dataGridView1.Columns[0].Width = 250;
                dataGridView1.Columns[1].Width = 220;
                con.Close();
                label23.Text = "ФИО по убыванию";
            }
            catch { }
        }

        private void radioButton7_CheckedChanged(object sender, EventArgs e)
        {
            try {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select  Код_Студента ,      Студенты.ФИО , Студенты.Год_Поступления , Студенты.Форма_обучения, Специальности.Специальности  from Студенты, Группы, Специальности where Специальности.Код_специальности =Группы.Код_специальности and Группы.Код_группы = Студенты.Код_группы ORDER BY Студенты.Год_Поступления";

                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);


                dataGridView1.DataSource = dt;
                dataGridView1.Columns[0].Width = 250;
                dataGridView1.Columns[1].Width = 220;
                con.Close();
                label23.Text = "Год по возрастанию";
            }
            catch { }
        }

        private void radioButton8_CheckedChanged(object sender, EventArgs e)
        {
            try {
                con.Open();
                OleDbCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "Select  Код_Студента ,      Студенты.ФИО , Студенты.Год_Поступления , Студенты.Форма_обучения, Специальности.Специальности  from Студенты, Группы, Специальности where Специальности.Код_специальности =Группы.Код_специальности and Группы.Код_группы = Студенты.Код_группы ORDER BY Студенты.Год_Поступления DESC";

                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);


                dataGridView1.DataSource = dt;
                dataGridView1.Columns[0].Width = 250;
                dataGridView1.Columns[1].Width = 220;
                con.Close();
                label23.Text = "Год по убыванию";
            }
            catch { }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            Form4 ee = new Form4();
            ee.Show();
            this.Close();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            try {
                string queryString5 = "SELECT Sum(Оценка) FROM Отметки where год= " + dateTimePicker6.Value.Date.Year + " and Отметки.Код_студента=" + (Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value)) + "";
                string queryString55 = "SELECT count(Код_успеваемости_студента) FROM Отметки where год=" + dateTimePicker6.Value.Date.Year + " and Код_студента=" + (Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value)) + " ";
           
                try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    OleDbCommand command55 = new OleDbCommand(queryString5, connection);
                    OleDbCommand command555 = new OleDbCommand(queryString55, connection);

                    
                    connection.Open();
                    ////2019
                    String ss = command55.ExecuteScalar().ToString();
                    String ss1 = command555.ExecuteScalar().ToString();
                    int a = Convert.ToInt32(ss);
                    int aa = Convert.ToInt32(ss1);
                    double q = 0;
                    q = a / aa;
                    label41.Text =Convert.ToString(q);

                }
            }
            catch (Exception ex)
            {
                label41.Text = "";
            }


            string queryString3 = "SELECT Sum(Оценка) FROM Отметки where год=" + dateTimePicker6.Value.Date.Year + "";
            string queryString33 = "SELECT count(Код_успеваемости_студента) FROM Отметки where год=" + dateTimePicker6.Value.Date.Year + "";
            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    OleDbCommand command = new OleDbCommand(queryString3, connection);
                    OleDbCommand command1 = new OleDbCommand(queryString33, connection);


                    connection.Open();
                    ////2019
                    String ss = command.ExecuteScalar().ToString();
                    String ss1 = command1.ExecuteScalar().ToString();
                    int a = Convert.ToInt32(ss);
                    int aa = Convert.ToInt32(ss1);
                    int q = 0;
                    q = a / aa;
                    label42.Text ="Средний бал среди студентов в"+ dateTimePicker6.Value.Date.Year + " году равен " +Convert.ToString(q);
                }

            }
            catch (Exception ex)
            {
                label42.Text = "";
            }
            }
            catch { }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            Help.ShowHelp(this, @"NewProject.chm", HelpNavigator.TableOfContents);
        }
    }

}
