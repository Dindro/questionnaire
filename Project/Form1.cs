using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace Project
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            Data.form = this;
            Data.panel = this.panel1;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Data.OleDbConnection.Close();
            Data.OleDbConnection.Dispose();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (!File.Exists("DataBaseFinal.mdb"))
                File.WriteAllBytes("DataBaseFinal.mdb", Properties.Resources.DataBaseFinal);
            if (!File.Exists("Microsoft.Office.Interop.Word.dll"))
                File.WriteAllBytes("Microsoft.Office.Interop.Word.dll", Properties.Resources.Microsoft_Office_Interop_Word);
            if (!File.Exists("Office.dll"))
                File.WriteAllBytes("Office.dll", Properties.Resources.Office);
            if (!File.Exists("System.Data.dll"))
                File.WriteAllBytes("System.Data.dll", Properties.Resources.System_Data);

            string connectionString ="provider=Microsoft.Jet.OLEDB.4.0;data source=" +System.IO.Path.Combine(Application.StartupPath, "DataBaseFinal.mdb");
            OleDbConnection myOleDbConnection = new OleDbConnection(connectionString);
            Data.OleDbConnection = myOleDbConnection;
            Data.OleDbConnection.Open();


            if (!panel1.Controls.Contains(Start.Instance)) //если нет в массиве то добавляем
            {
                panel1.Controls.Add(Start.Instance);
                Start.Instance.Dock = DockStyle.Fill;
                Start.Instance.BringToFront();
            }
            else
                Start.Instance.BringToFront();


        }
    }
}
