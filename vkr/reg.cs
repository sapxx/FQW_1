using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace vkr
{
    
    public partial class reg : Form
    {
        
        public string log;
        public static string connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=бд.mdb;";
        private OleDbConnection myConnection;
        public reg()
        {
            InitializeComponent();
            myConnection = new OleDbConnection(connectString);
            
        }
        private void reg_Load(object sender, EventArgs e)
        {
            panel2.Hide();
            panel1.Hide();
            login.Clear();
            password.Clear();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            
            panel2.Show();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            panel1.Show();
            panel2.Hide();
        }

        public void enter_Click(object sender, EventArgs e)
        {
            
            log = login.Text;
            myConnection.Close();
            if (login.Text == "admin")
            {
                if (password.Text == "admin")
                {
                    
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
                else
                {
                    MessageBox.Show("Не верный пароль");
                }
            }
            if (login.Text == "qq")
            {
                if (password.Text == "qq")
                {
                    
                    this.DialogResult = DialogResult.OK;
                    
                    this.Close();
                }
                else
                {
                    MessageBox.Show("Не верный пароль");
                }

            }

            if (login.Text == null | password.Text == null)
            {
                MessageBox.Show("Не верный логин либо пароль");
            }
        }

        private void reg_FormClosed(object sender, FormClosedEventArgs e)
        {

            if (login.Text != "admin" | password.Text != "admin")
                if (login.Text != "qq" | password.Text != "qq")
                {
                    Application.Exit();

                }
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            panel2.Hide();
            panel1.Hide();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            panel2.Hide();
            panel1.Hide();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            panel1.Hide();
            panel2.Show();
            log = null;
        }

        private void input_Click(object sender, EventArgs e)
        {
            myConnection.Open();
            string name = richINC.Text;
            string trouble = richTrouble.Text;
            string info = richinfo.Text;
            string date = dateTimePicker1.Text;
            string worker = "Новая";
            string adres = richTextBox2.Text;
            string phone = richTextBox1.Text;
            // cmd = new SqlCommand("Insert Into into Values ('" + richTextBox1.Text + "', '" + richTextBox2.Text + "', '" + dateTimePicker1.Value.ToString("dd-MM-yyyy") + "')", conn); // создаем SQL запрос
            // текст запроса
            string query = "INSERT INTO TableNF(Организация,Проблема,Оборудывание,Дата_регистрации,Статус,Адрес,Телефон_Email)"
                                         + "VALUES ('" + name + "','" + trouble + "','" + info + "','" + date + "','" + worker + "','" + adres + "','" + phone + "')";
            // создаем объект OleDbCommand для выполнения запроса к БД MS Access
            OleDbCommand command = new OleDbCommand(query, myConnection);

            // выполняем запрос к MS Access
            command.ExecuteNonQuery();
            MessageBox.Show("Данные переданны");
            richINC.Clear();
            richTrouble.Clear();
            richinfo.Clear();
            dateTimePicker1.ResetText();
            panel1.Hide();
            panel2.Hide();
            myConnection.Close();
        }
    }
}
