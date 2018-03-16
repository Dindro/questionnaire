using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace Project
{
    public partial class Tables : UserControl
    {
        private static Tables instance;
        public static Tables Instance
        {
            get
            {
                if (instance == null)
                    instance = new Tables();
                return instance;
            }
        }

        public Tables()
        {
            InitializeComponent();
            Data.form.Text = "Итог";
        }

        private void Tables_Load(object sender, EventArgs e)
        {
            Data.SetInObjectFromDataBaseForTable("Education", 1 ,dataGridView1);
            Data.SetInObjectFromDataBaseForTable("Scientific", 2, dataGridView2);
            Data.SetInObjectFromDataBaseForTable("Culture", 3, dataGridView3);
            Data.SetInObjectFromDataBaseForTable("Publics", 4, dataGridView4);
            Data.SetInObjectFromDataBaseForTable("Sports", 5, dataGridView5);
        }

        private void downloadA_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount == 0)
                return;
            int row = dataGridView1.CurrentRow.Index;
            int id = Convert.ToInt32(dataGridView1[0, row].Value);
            Human human = new Human();

            human.SetInObjectFromDataBase(id);
            human.SetInWord();
        }

        private void downloadD_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount == 0)
                return;
            int row = dataGridView1.CurrentRow.Index;
            int id = Convert.ToInt32(dataGridView1[0, row].Value);
            Education educationActivity = new Education(new double[] { 2, 3, 4, 5, 6 },
                                                        new double[] { 2, 3, 4, 5, 6 });
            educationActivity.SetInObjectFromDataBase(id);
            educationActivity.SetInWord();
        }
        //
        private void downloadA2_Click(object sender, EventArgs e)
        {
            if (dataGridView2.RowCount == 0)
                return;
            int row = dataGridView2.CurrentRow.Index;
            int id = Convert.ToInt32(dataGridView2[0, row].Value);
            Human human = new Human();

            human.SetInObjectFromDataBase(id);
            human.SetInWord();
        }

        private void downloadD2_Click(object sender, EventArgs e)
        {
            if (dataGridView2.RowCount == 0)
                return;
            int row = dataGridView2.CurrentRow.Index;
            int id = Convert.ToInt32(dataGridView2[0, row].Value);
            ScientificResearch scientificResearch = new ScientificResearch(new double[] { 2, 3, 4, 5, 6 },
                                                       new double[] { 8 },
                                                       new double[] { 8 },
                                                        new double[] { 2, 3, 4, 5, 6 });
            scientificResearch.SetInObjectFromDataBase(id);
            scientificResearch.SetInWord();
        }
        //
        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView3.RowCount == 0)
                return;
            int row = dataGridView3.CurrentRow.Index;
            int id = Convert.ToInt32(dataGridView3[0, row].Value);
            Human human = new Human();

            human.SetInObjectFromDataBase(id);
            human.SetInWord();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (dataGridView3.RowCount == 0)
                return;
            int row = dataGridView3.CurrentRow.Index;
            int id = Convert.ToInt32(dataGridView3[0, row].Value);
            
            CulturalСreative culturalСreative = new CulturalСreative(new double[] { 2, 3, 4, 5, 6 },
                                                    new double[] { 2, 3, 4, 5, 6 },
                                                    new double[] { 8 });
            culturalСreative.SetInObjectFromDataBase(id);
            culturalСreative.SetInWord();
        }
        //
        private void button4_Click(object sender, EventArgs e)
        {
            if (dataGridView4.RowCount == 0)
                return;
            int row = dataGridView4.CurrentRow.Index;
            int id = Convert.ToInt32(dataGridView4[0, row].Value);
            Human human = new Human();

            human.SetInObjectFromDataBase(id);
            human.SetInWord();

        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (dataGridView4.RowCount == 0)
                return;
            int row = dataGridView4.CurrentRow.Index;
            int id = Convert.ToInt32(dataGridView4[0, row].Value);
            Public publics = new Public(new double[] { 2, 3, 4, 5, 6 },
                                 new double[] { 2, 3, 4, 5, 6 });
            publics.SetInObjectFromDataBase(id);
            publics.SetInWord();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (dataGridView5.RowCount == 0)
                return;
            int row = dataGridView5.CurrentRow.Index;
            int id = Convert.ToInt32(dataGridView5[0, row].Value);
            Human human = new Human();

            human.SetInObjectFromDataBase(id);
            human.SetInWord();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (dataGridView5.RowCount == 0)
                return;
            int row = dataGridView5.CurrentRow.Index;
            int id = Convert.ToInt32(dataGridView5[0, row].Value);
            Sport sport = new Sport(new double[] { 2, 3, 4, 5, 6 },
                              new double[] { 2, 3, 4, 5, 6 },
                              6);
            sport.SetInObjectFromDataBase(id);
            sport.SetInWord();
        }

        private void delete_Click(object sender, EventArgs e)
        {
            if (dataGridView2.RowCount == 0)
                return;
            int row = dataGridView2.CurrentRow.Index;
            int id = Convert.ToInt32(dataGridView2[0, row].Value);
            ScientificResearch.DeleteDataBase(id);
            dataGridView2.Rows.RemoveAt(row);
        }

        private void delete1_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount == 0)
                return;
            int row = dataGridView1.CurrentRow.Index;
            int id = Convert.ToInt32(dataGridView1[0, row].Value);
            Education.DeleteDataBase(id);
            dataGridView1.Rows.RemoveAt(row);
        }

        private void delete3_Click(object sender, EventArgs e)
        {
            if (dataGridView3.RowCount == 0)
                return;
            int row = dataGridView3.CurrentRow.Index;
            int id = Convert.ToInt32(dataGridView3[0, row].Value);
            CulturalСreative.DeleteDataBase(id);
            dataGridView3.Rows.RemoveAt(row);
        }

        private void delete4_Click(object sender, EventArgs e)
        {
            if (dataGridView4.RowCount == 0)
                return;
            int row = dataGridView4.CurrentRow.Index;
            int id = Convert.ToInt32(dataGridView4[0, row].Value);
            Public.DeleteDataBase(id);
            dataGridView4.Rows.RemoveAt(row);
        }

        private void delete5_Click(object sender, EventArgs e)
        {
            if (dataGridView5.RowCount == 0)
                return;
            int row = dataGridView5.CurrentRow.Index;
            int id = Convert.ToInt32(dataGridView5[0, row].Value);
            Sport.DeleteDataBase(id);
            dataGridView5.Rows.RemoveAt(row);
        }
    }
}
