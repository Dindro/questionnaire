using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;

namespace Project
{
    public partial class Start : UserControl
    {
        private static Start instance;
        public static Start Instance
        {
            get
            {
                if (instance == null)
                    instance = new Start();
                return instance;
            }
        }

        public Start()
        {
            InitializeComponent();
        }

        private void registration_Click(object sender, EventArgs e)
        {
            if (!Data.panel.Controls.Contains(Worksheet.Instance)) //если нет в массиве то добавляем
            {
                Data.panel.Controls.Add(Worksheet.Instance);
                Worksheet.Instance.Dock = DockStyle.Fill;
                Worksheet.Instance.BringToFront();
            }
            else
                Worksheet.Instance.BringToFront();
        }

        private void enter_Click(object sender, EventArgs e)
        {
            if (!Data.panel.Controls.Contains(Tables.Instance))
            {
                Data.panel.Controls.Add(Tables.Instance);
                Tables.Instance.Dock = DockStyle.Fill;
                Tables.Instance.BringToFront();
            }
            else
                Tables.Instance.BringToFront();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            Password pas = new Password();
            pas.ShowDialog();
        }
    }
}
