using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace Project
{
    public partial class Password : Form
    {
        public Password()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
            this.Dispose();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Login login = new Login();
            if (login.Check(textBox1.Text))
            {
                this.Close();
                this.Dispose();
            }
            else
                label1.Text = "Не правильно";
        }
    }
}
