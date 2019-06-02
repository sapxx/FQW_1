using System;
using System.Threading;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing.Printing;
using System.Drawing;


namespace vkr
{
    public partial class Form1 : Form
    {
        public static string connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=бд.mdb;";
        private OleDbConnection myConnection;
        reg regf = new reg();
        public Form1()
        {
            InitializeComponent();
            myConnection = new OleDbConnection(connectString);
            myConnection.Open();
            regf.ShowDialog();
            if (regf.DialogResult == DialogResult.OK)
            {
                this.Show();
            }
            
            

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "бдDataSet3.TableWorkers". При необходимости она может быть перемещена или удалена.
            this.tableWorkersTableAdapter.Fill(this.бдDataSet3.TableWorkers);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "бдDataSet.TableNF1". При необходимости она может быть перемещена или удалена.
            this.tableNF1TableAdapter.Fill(this.бдDataSet.TableNF1);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "бдDataSet2.TableAF". При необходимости она может быть перемещена или удалена.
            this.tableAFTableAdapter.Fill(this.бдDataSet2.TableAF);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "бдDataSet1.TableNF". При необходимости она может быть перемещена или удалена.
            this.tableNFTableAdapter.Fill(this.бдDataSet1.TableNF);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "бдDataSet.TableWorkers". При необходимости она может быть перемещена или удалена.
            this.tableWorkersTableAdapter.Fill(this.бдDataSet.TableWorkers);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "бдDataSet.TableInc". При необходимости она может быть перемещена или удалена.
            this.tableIncTableAdapter.Fill(this.бдDataSet.TableInc);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "бдDataSet.TableAF". При необходимости она может быть перемещена или удалена.
            this.tableAFTableAdapter.Fill(this.бдDataSet.TableAF);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "бдDataSet.TableNF". При необходимости она может быть перемещена или удалена.
            this.tableNFTableAdapter.Fill(this.бдDataSet.TableNF);
            label1.Text = regf.log;
            if (label1.Text == "admin")
            {
            }
            else
            {
                tabControl1.TabPages.Remove(tabPage4);
            }

            


        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.tableWorkersTableAdapter.Update(this.бдDataSet.TableWorkers);
            this.tableIncTableAdapter.Update(this.бдDataSet.TableInc);
            this.tableAFTableAdapter.Update(this.бдDataSet.TableAF);
            this.tableNFTableAdapter.Update(this.бдDataSet.TableNF);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.бдDataSet.Clear();
            this.tableWorkersTableAdapter.Fill(this.бдDataSet.TableWorkers);
            this.tableIncTableAdapter.Fill(this.бдDataSet.TableInc);
            this.tableAFTableAdapter.Fill(this.бдDataSet.TableAF);
            this.tableNFTableAdapter.Fill(this.бдDataSet.TableNF);
            this.tableNF1TableAdapter.Fill(this.бдDataSet.TableNF1);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            int a = dataGridView1.CurrentRow.Index;
            dataGridView1.Rows.Remove(dataGridView1.Rows[a]);
            dataGridView1.Refresh();
            #region save&update
            this.tableWorkersTableAdapter.Update(this.бдDataSet.TableWorkers);
            this.tableIncTableAdapter.Update(this.бдDataSet.TableInc);
            this.tableAFTableAdapter.Update(this.бдDataSet.TableAF);
            this.tableNFTableAdapter.Update(this.бдDataSet.TableNF);
            this.бдDataSet.Clear();
            this.tableWorkersTableAdapter.Fill(this.бдDataSet.TableWorkers);
            this.tableIncTableAdapter.Fill(this.бдDataSet.TableInc);
            this.tableAFTableAdapter.Fill(this.бдDataSet.TableAF);
            this.tableNFTableAdapter.Fill(this.бдDataSet.TableNF);
            this.tableNF1TableAdapter.Fill(this.бдDataSet.TableNF1);
            #endregion



        }
        void func(string param)
        {
            // обработка
        }
        private void button6_Click(object sender, EventArgs e)
        {
            using (add fMyForm = new add())                
            {
                if (fMyForm.ShowDialog() == DialogResult.OK)
                {
                    string fio = fMyForm.flo;
                    string query = "INSERT INTO TableWorkers(ФИО)"
                                         + "VALUES ('" + fio + "')";
                    OleDbCommand command = new OleDbCommand(query, myConnection);
                    command.ExecuteNonQuery();

                    this.tableWorkersTableAdapter.Update(this.бдDataSet.TableWorkers);
                    this.tableIncTableAdapter.Update(this.бдDataSet.TableInc);
                    this.tableAFTableAdapter.Update(this.бдDataSet.TableAF);
                    this.tableNFTableAdapter.Update(this.бдDataSet.TableNF);
                    Thread.Sleep(1000);
                    this.button2.PerformClick();
                }

            }

        }

        private void button7_Click(object sender, EventArgs e)
        {
            using (incadd y1Form = new incadd())

            {
                if (y1Form.ShowDialog() == DialogResult.OK)
                {
                    string inca = y1Form.inca;
                    string adress = y1Form.adra;
                    string phone = y1Form.phoa + ";" + y1Form.ema;
                    string query = "INSERT INTO TableInc(Организация,Адрес,Телефон_Email)"
                                         + "VALUES ('" + inca + "','" + adress + "','" + phone + "')";
                    OleDbCommand command = new OleDbCommand(query, myConnection);
                    command.ExecuteNonQuery();
                    this.tableWorkersTableAdapter.Update(this.бдDataSet.TableWorkers);
                    this.tableIncTableAdapter.Update(this.бдDataSet.TableInc);
                    this.tableAFTableAdapter.Update(this.бдDataSet.TableAF);
                    this.tableNFTableAdapter.Update(this.бдDataSet.TableNF);
                    Thread.Sleep(1000);
                    this.button2.PerformClick();
                }

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            int a = dataGridView1.CurrentRow.Index;
            string num = dataGridView1.Rows[a].Cells[0].Value.ToString();
            string date = dateTimePicker1.Text;
            string inc = dataGridView1.Rows[a].Cells[1].Value.ToString();
            string info = dataGridView1.Rows[a].Cells[7].Value.ToString();
            string work = dataGridView1.Rows[a].Cells[10].Value.ToString();
            string adres = dataGridView1.Rows[a].Cells[2].Value.ToString();
            string phone = dataGridView1.Rows[a].Cells[3].Value.ToString();
            string data2 = dataGridView1.Rows[a].Cells[4].Value.ToString();
            string data3 = dataGridView1.Rows[a].Cells[5].Value.ToString();
            string summ = dataGridView1.Rows[a].Cells[8].Value.ToString();
            string stat = "Закрыто";
            string worker = dataGridView1.Rows[a].Cells[9].Value.ToString();

            string query = "INSERT INTO TableAF(Организация,Дата_регистрации,Дата_закрытия,Статус,Оборудывание,Сумма,Исполнитель,Выполнено_работ)"
                                         + "VALUES ('" + inc + "','" + data2 + "','" + data3 + "','" + stat + "','" + info + "','" + summ + "','" + worker + "','" + work + "')";
            // создаем объект OleDbCommand для выполнения запроса к БД MS Access
            OleDbCommand command = new OleDbCommand(query, myConnection);

            // выполняем запрос к MS Access
            command.ExecuteNonQuery();
            dataGridView1.Rows.Remove(dataGridView1.Rows[a]);
            this.tableNFTableAdapter.Update(this.бдDataSet.TableNF);
            MessageBox.Show("Статус заявки переведен на 'выполнена'. Заявка отправлена в архив. Дождитесь открытия отчета Word");
            //--------------------------------------------------------
            //--------------------------------------------------------
            //--------------------------------------------------------
            //--------------------------------------------------------
            //Создаём новый Word.Application
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            string path = Environment.CurrentDirectory + "/akt.dot";
            //Загружаем документ
            Microsoft.Office.Interop.Word.Document doc = null;
            object fileName = path;
            object falseValue = false;
            object trueValue = true;
            object missing = Type.Missing;

            doc = app.Documents.Open(ref fileName, ref missing, ref trueValue,
            ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing);

            //Теперь у нас есть документ который мы будем менять.

            //Очищаем параметры поиска
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();
            //--------------------------------------------------------
            //--------------------------------------------------------
            //--------------------------------------------------------
            //--------------------------------------------------------
            //Задаём параметры замены и выполняем замену.
            object findText = "<date>";
            object replaceWith = date;
            object replace = 2;
            app.Selection.Find.Execute(ref findText, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith,
            ref replace, ref missing, ref missing, ref missing, ref missing);

            ///////////////////
            ///
            object findText1 = "<num>";
            object replaceWith1 = num;
            object replace1 = 2;
            app.Selection.Find.Execute(ref findText1, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith1,
            ref replace1, ref missing, ref missing, ref missing, ref missing);
            //////////////////
            ///
            object findText2 = "<inc>";
            object replaceWith2 = inc;
            object replace2 = 2;
            app.Selection.Find.Execute(ref findText2, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith2,
            ref replace2, ref missing, ref missing, ref missing, ref missing);

            ///////////////////////////

            object findText4 = "<info>";
            object replaceWith4 = info;
            object replace4 = 2;
            app.Selection.Find.Execute(ref findText4, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith4,
            ref replace4, ref missing, ref missing, ref missing, ref missing);
            ///
            object findText5 = "<work>";
            object replaceWith5 = work;
            object replace5 = 2;
            app.Selection.Find.Execute(ref findText5, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith5,
            ref replace5, ref missing, ref missing, ref missing, ref missing);
            //
            object findText6 = "<street>";
            object replaceWith6 = adres;
            object replace6 = 2;
            app.Selection.Find.Execute(ref findText6, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith6,
            ref replace6, ref missing, ref missing, ref missing, ref missing);
            //
            object findText7 = "<phone>";
            object replaceWith7 = phone;
            object replace7 = 2;
            app.Selection.Find.Execute(ref findText7, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith7,
            ref replace7, ref missing, ref missing, ref missing, ref missing);

            // Открываем документ для просмотра.
            app.Visible = true;
        }

        private void tableNFBindingSource_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void fill1ToolStripButton_Click(object sender, EventArgs e)
        {
            try
            {
                this.tableNFTableAdapter.Fill1(this.бдDataSet.TableNF);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }

        private void button5_Click(object sender, EventArgs e)
        {
            
                string a = comboBox1.Text;
                string query = "SELECT * From TableAF WHERE Исполнитель = '" + a + "'";
                // создаем объект OleDbCommand для выполнения запроса к БД MS Access
                OleDbCommand command = new OleDbCommand(query, myConnection);

                // выполняем запрос к MS Access
                command.ExecuteNonQuery();
                DataTable dt = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(command);
                da.Fill(dt);
                dataGridView2.DataSource = dt;
                

            
        }

        private void button8_Click(object sender, EventArgs e)
        {
            dataGridView2.DataSource = tableAFBindingSource;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            
        }

        private void button9_Click_1(object sender, EventArgs e)
        {
            this.Hide();
            regf.ShowDialog();
            if (regf.DialogResult == DialogResult.OK)
            {
                this.Show();
                label1.Text = regf.log;
                if (label1.Text == "admin")
                {
                    this.Refresh();
                    this.tabPage4 = new System.Windows.Forms.TabPage();
                    this.tabPage4.SuspendLayout();
                    this.tabControl1.Controls.Add(this.tabPage4);
                    // 
                    // tabPage4
                    // 
                    this.tabPage4.Controls.Add(this.button6);
                    this.tabPage4.Controls.Add(this.dataGridView4);
                    this.tabPage4.Location = new System.Drawing.Point(4, 22);
                    this.tabPage4.Name = "tabPage4";
                    this.tabPage4.Size = new System.Drawing.Size(1040, 384);
                    this.tabPage4.TabIndex = 3;
                    this.tabPage4.Text = "Инженеры";
                    this.tabPage4.UseVisualStyleBackColor = true;
                    this.tabPage4.ResumeLayout(false);
                    button2.PerformClick();
                }
                else
                {
                    tabControl1.TabPages.Remove(tabPage4);
                }
                
            }

        }
    }
}
