using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace vkr
{
    public partial class incadd : Form
    {
        public string inca = "Ошибка";
        public string adra = "Ошибка";
        public string phoa = "Ошибка";
        public string ema = "Ошибка";
        public incadd()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
            inca = textBox1.Text;
            adra = textBox2.Text;
            phoa = textBox3.Text;
            ema = textBox4.Text;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
