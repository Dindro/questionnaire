using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
//using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;
using System.Data.OleDb;

namespace Project
{
    public partial class ScientificResearchAcivity : UserControl
    {
        private static ScientificResearchAcivity instance;
        public static ScientificResearchAcivity Instance
        {
            get
            {
                if (instance == null)
                    instance = new ScientificResearchAcivity();
                return instance;
            }
        }

        ScientificResearch scientificResearch;
        public DataGridView[] tables;
        public DataGridView[] tablesA;
        public DataGridView[] tablesB;
        public DataGridView[] tablesC;
        public List<Control> fields;

        public ScientificResearchAcivity()
        {
            InitializeComponent();
            Data.form.Text = "Научно-исследовательская деятельность";
            panelMain.AutoScroll = true;
            scientificResearch = new ScientificResearch(new double[] { 2, 3, 4, 5, 6 },
                                                        new double[] { 8 },
                                                        new double[] { 8 },
                                                        new double[] { 2, 3, 4, 5, 6 });
            fields = new List<Control>() { maskedTextBoxPeriod};
            tables = new DataGridView[] { tableE1,tableD1,
                                           tableE2,tableD2,
                                           tableE3,tableD3,
                                           tableE4,tableD4,
                                           tableE5,tableD5};
            tablesA = new DataGridView[] { tableAE1,tableAD1,
                                           tableAE2,tableAD2,
                                           tableAE3,tableAD3,
                                           tableAE4,tableAD4,
                                           tableAE5,tableAD5};
            tablesB = new DataGridView[] { tableBE1, tableBD1 };
            tablesC = new DataGridView[] { tableCE1, tableCD1 };
            fields.AddRange(tables);
            fields.AddRange(tablesA);
            fields.AddRange(tablesB);
            fields.AddRange(tablesC);
            result.Text = "0";
            radioButton1And1And1.Checked = true;

        }

        #region Пункт #2.1
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1And1And1.Checked)
            {
                scientificResearch.Status = 1;
                ocenka1.Text = scientificResearch.progress.ToString();
                result.Text = scientificResearch.DecisionResult().ToString();
            }
            else
                ocenka1.Text = "";
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1And1And2.Checked)
            {
                scientificResearch.Status = 2;
                ocenka2.Text = scientificResearch.progress.ToString();
                result.Text = scientificResearch.DecisionResult().ToString();
            }
            else
                ocenka2.Text = "";
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1And1And3.Checked)
            {
                scientificResearch.Status = 3;
                ocenka3.Text = scientificResearch.progress.ToString();
                result.Text = scientificResearch.DecisionResult().ToString();
            }
            else
                ocenka3.Text = "";
        }
        #endregion

        #region Пункт #2.2
        private void plus1_Click(object sender, EventArgs e)
        {
            Activity activity = scientificResearch;
            activity.Plus(tableE1, tableD1);
            ball1.Text = scientificResearch.levels[0].DecisionResult(tableE1.RowCount);
            result.Text = scientificResearch.DecisionResult().ToString();
            activity = null;
        }

        private void minus1_Click(object sender, EventArgs e)
        {
            Activity activity = scientificResearch;
            activity.Minus(tableE1, tableD1); //удаляем мероприятие из таблицы
            ball1.Text = scientificResearch.levels[0].DecisionResult(tableE1.RowCount);  //вычесляем результат для бокового
            result.Text = scientificResearch.DecisionResult().ToString(); //вычисляем результат общий
            activity = null;
        }

        private void plus2_Click(object sender, EventArgs e)
        {
            Activity activity = scientificResearch;
            activity.Plus(tableE2, tableD2);
            ball2.Text = scientificResearch.levels[1].DecisionResult(tableE2.RowCount);
            result.Text = scientificResearch.DecisionResult().ToString();
            activity = null;
        }

        private void minus2_Click(object sender, EventArgs e)
        {
            Activity activity = scientificResearch;
            activity.Minus(tableE2, tableD2); //удаляем мероприятие из таблицы
            ball2.Text = scientificResearch.levels[1].DecisionResult(tableE2.RowCount);  //вычесляем результат для бокового
            result.Text = scientificResearch.DecisionResult().ToString(); //вычисляем результат общий
            activity = null;
        }

        private void plus3_Click(object sender, EventArgs e)
        {
            Activity activity = scientificResearch;
            activity.Plus(tableE3, tableD3);
            ball3.Text = scientificResearch.levels[2].DecisionResult(tableE3.RowCount);
            result.Text = scientificResearch.DecisionResult().ToString();
            activity = null;
        }

        private void minus3_Click(object sender, EventArgs e)
        {
            Activity activity = scientificResearch;
            activity.Minus(tableE3, tableD3); //удаляем мероприятие из таблицы
            ball3.Text = scientificResearch.levels[2].DecisionResult(tableE3.RowCount);  //вычесляем результат для бокового
            result.Text = scientificResearch.DecisionResult().ToString(); //вычисляем результат общий
            activity = null;
        }

        private void plus4_Click(object sender, EventArgs e)
        {
            Activity activity = scientificResearch;
            activity.Plus(tableE4, tableD4);
            ball4.Text = scientificResearch.levels[3].DecisionResult(tableE4.RowCount);
            result.Text = scientificResearch.DecisionResult().ToString();
            activity = null;
        }

        private void minus4_Click(object sender, EventArgs e)
        {
            Activity activity = scientificResearch;
            activity.Minus(tableE4, tableD4); //удаляем мероприятие из таблицы
            ball4.Text = scientificResearch.levels[3].DecisionResult(tableE4.RowCount);  //вычесляем результат для бокового
            result.Text = scientificResearch.DecisionResult().ToString(); //вычисляем результат общий
            activity = null;
        }

        private void plus5_Click(object sender, EventArgs e)
        {
            Activity activity = scientificResearch;
            activity.Plus(tableE5, tableD5);
            ball5.Text = scientificResearch.levels[4].DecisionResult(tableE5.RowCount);
            result.Text = scientificResearch.DecisionResult().ToString();
            activity = null;
        }

        private void minus5_Click(object sender, EventArgs e)
        {
            Activity activity = scientificResearch;
            activity.Minus(tableE5, tableD5); //удаляем мероприятие из таблицы
            ball5.Text = scientificResearch.levels[4].DecisionResult(tableE5.RowCount);  //вычесляем результат для бокового
            result.Text = scientificResearch.DecisionResult().ToString(); //вычисляем результат общий
            activity = null;
        }
        #endregion

        #region Пункт #5.3
        private void plusA1_Click(object sender, EventArgs e)
        {
            scientificResearch.Plus(tableAE1, tableAD1);
            ballA1.Text = scientificResearch.performances[0].DecisionResult(tableAE1.RowCount);
            result.Text = scientificResearch.DecisionResult().ToString();
        }

        private void minusA1_Click(object sender, EventArgs e)
        {
            scientificResearch.Minus(tableAE1, tableAD1); //удаляем мероприятие из таблицы
            ballA1.Text = scientificResearch.performances[0].DecisionResult(tableAE1.RowCount);  //вычесляем результат для бокового
            result.Text = scientificResearch.DecisionResult().ToString(); //вычисляем результат общий
        }

        private void plusA2_Click(object sender, EventArgs e)
        {
            scientificResearch.Plus(tableAE2, tableAD2);
            ballA2.Text = scientificResearch.performances[1].DecisionResult(tableAE2.RowCount);
            result.Text = scientificResearch.DecisionResult().ToString();
        }

        private void minusA2_Click(object sender, EventArgs e)
        {
            scientificResearch.Minus(tableAE2, tableAD2); //удаляем мероприятие из таблицы
            ballA2.Text = scientificResearch.performances[1].DecisionResult(tableAE2.RowCount);  //вычесляем результат для бокового
            result.Text = scientificResearch.DecisionResult().ToString(); //вычисляем результат общий
        }

        private void plusA3_Click(object sender, EventArgs e)
        {
            scientificResearch.Plus(tableAE3, tableAD3);
            ballA3.Text = scientificResearch.performances[2].DecisionResult(tableAE3.RowCount);
            result.Text = scientificResearch.DecisionResult().ToString();
        }

        private void minusA3_Click(object sender, EventArgs e)
        {
            scientificResearch.Minus(tableAE3, tableAD3); //удаляем мероприятие из таблицы
            ballA3.Text = scientificResearch.performances[2].DecisionResult(tableAE3.RowCount);  //вычесляем результат для бокового
            result.Text = scientificResearch.DecisionResult().ToString(); //вычисляем результат общий
        }

        private void plusA4_Click(object sender, EventArgs e)
        {
            scientificResearch.Plus(tableAE4, tableAD4);
            ballA4.Text = scientificResearch.performances[3].DecisionResult(tableAE4.RowCount);
            result.Text = scientificResearch.DecisionResult().ToString();
        }

        private void minusA4_Click(object sender, EventArgs e)
        {
            scientificResearch.Minus(tableAE4, tableAD4); //удаляем мероприятие из таблицы
            ballA4.Text = scientificResearch.performances[3].DecisionResult(tableAE4.RowCount);  //вычесляем результат для бокового
            result.Text = scientificResearch.DecisionResult().ToString(); //вычисляем результат общий
        }

        private void plusA5_Click(object sender, EventArgs e)
        {
            scientificResearch.Plus(tableAE5, tableAD5);
            ballA5.Text = scientificResearch.performances[4].DecisionResult(tableAE5.RowCount);
            result.Text = scientificResearch.DecisionResult().ToString();
        }

        private void minusA5_Click(object sender, EventArgs e)
        {
            scientificResearch.Minus(tableAE5, tableAD5); //удаляем мероприятие из таблицы
            ballA5.Text = scientificResearch.performances[4].DecisionResult(tableAE5.RowCount);  //вычесляем результат для бокового
            result.Text = scientificResearch.DecisionResult().ToString(); //вычисляем результат общий
        }
        #endregion

        private new void Enter(object sender, EventArgs e)
        {
            Data.Enter((Control)sender);
        }

        private void finish_Click(object sender, EventArgs e)
        {
            if (!Data.IsFilled(fields))
            {
                MessageBox.Show("Заполните пустые поля!", "Внимание");
                return;
            }

            scientificResearch.period = maskedTextBoxPeriod.Text;
            scientificResearch.levels = scientificResearch.SetFromTables(scientificResearch.levels, tables);

            scientificResearch.documents = scientificResearch.SetFromTables(scientificResearch.documents, tablesB);
            scientificResearch.grants = scientificResearch.SetFromTables(scientificResearch.grants, tablesC);
            
            scientificResearch.performances = scientificResearch.SetFromTables(scientificResearch.performances, tablesA);

            OleDbCommand command = new OleDbCommand("INSERT INTO Scientific (Период, Статус, Уровень1,Уровень2,Уровень3,Уровень4,Уровень5, УровеньВ1, УровеньС1, УровеньА1, УровеньА2, УровеньА3, УровеньА4, УровеньА5, Результат)" +
                "VALUES(@Период, @Статус, @Уровень1,@Уровень2,@Уровень3,@Уровень4,@Уровень5, @УровеньВ1, @УровеньС1, @УровеньА1, @УровеньА2, @УровеньА3, @УровеньА4, @УровеньА5, @Результат)", Data.OleDbConnection);
            command.Parameters.AddWithValue("Период", scientificResearch.period.ToString());
            command.Parameters.AddWithValue("Статус", scientificResearch.status.ToString());

            command.Parameters.AddWithValue("Уровень1", scientificResearch.levels[0].SetEventsInDataBase());
            command.Parameters.AddWithValue("Уровень2", scientificResearch.levels[1].SetEventsInDataBase());
            command.Parameters.AddWithValue("Уровень3", scientificResearch.levels[2].SetEventsInDataBase());
            command.Parameters.AddWithValue("Уровень4", scientificResearch.levels[3].SetEventsInDataBase());
            command.Parameters.AddWithValue("Уровень5", scientificResearch.levels[4].SetEventsInDataBase());

            command.Parameters.AddWithValue("УровеньВ1", scientificResearch.documents[0].SetEventsInDataBase());
            command.Parameters.AddWithValue("УровеньС1", scientificResearch.grants[0].SetEventsInDataBase());

            command.Parameters.AddWithValue("УровеньА1", scientificResearch.performances[0].SetEventsInDataBase());
            command.Parameters.AddWithValue("УровеньА2", scientificResearch.performances[1].SetEventsInDataBase());
            command.Parameters.AddWithValue("УровеньА3", scientificResearch.performances[2].SetEventsInDataBase());
            command.Parameters.AddWithValue("УровеньА4", scientificResearch.performances[3].SetEventsInDataBase());
            command.Parameters.AddWithValue("УровеньА5", scientificResearch.performances[4].SetEventsInDataBase());

            command.Parameters.AddWithValue("Результат", scientificResearch.result.ToString());

            command.ExecuteNonQuery();

            int id = 0;
            command = new OleDbCommand("SELECT @@IDENTITY AS id", Data.OleDbConnection);
            OleDbDataReader sqlReaderA = null;
            sqlReaderA = command.ExecuteReader();
            while (sqlReaderA.Read())
            {
                id = Convert.ToInt32(sqlReaderA["id"]);
            }
            if (sqlReaderA != null)
                sqlReaderA.Close();

            scientificResearch.human.SetInDataBase(id);

            scientificResearch.SetInWord();

            if (!Data.panel.Controls.Contains(Tables.Instance)) //если нет в массиве то добавляем
            {
                Data.panel.Controls.Add(Tables.Instance);
                Tables.Instance.Dock = DockStyle.Fill;
                Tables.Instance.BringToFront();
            }
            else
                Tables.Instance.BringToFront();
        }

        //private void documnetCheck_CheckedChanged(object sender, EventArgs e)
        //{
        //    if (documnetCheck.Checked)
        //    {
        //        documentName.Enabled = true;
        //        documentDate.Enabled = true;
        //        scientificResearch.document.AreIs = 1;
        //        ballB1.Text = scientificResearch.document.progress.ToString();
        //    }
        //    else
        //    {
        //        documentName.Enabled = false;
        //        documentDate.Enabled = false;
        //        scientificResearch.document.AreIs = 0;
        //        ballB1.Text = "";
        //    }
        //    documentName.BackColor = Color.Empty;
        //    documentDate.BackColor = Color.Empty;
        //    result.Text = scientificResearch.DecisionResult().ToString();
        //}

        

        private void plusB1_Click(object sender, EventArgs e)
        {
            scientificResearch.Plus(tableBE1, tableBD1);
            ballB1.Text = scientificResearch.documents[0].DecisionResult(tableBE1.RowCount);
            result.Text = scientificResearch.DecisionResult().ToString();
        }

        private void minusB1_Click(object sender, EventArgs e)
        {
            scientificResearch.Minus(tableBE1, tableBD1); //удаляем мероприятие из таблицы
            ballB1.Text = scientificResearch.documents[0].DecisionResult(tableBE1.RowCount);  //вычесляем результат для бокового
            result.Text = scientificResearch.DecisionResult().ToString(); //вычисляем результат общий
        }

        private void plusC1_Click(object sender, EventArgs e)
        {
            scientificResearch.Plus(tableCE1, tableCD1);
            ballC1.Text = scientificResearch.grants[0].DecisionResult(tableCE1.RowCount);
            result.Text = scientificResearch.DecisionResult().ToString();
        }

        private void minusC1_Click(object sender, EventArgs e)
        {
            scientificResearch.Minus(tableCE1, tableCD1); //удаляем мероприятие из таблицы
            ballC1.Text = scientificResearch.grants[0].DecisionResult(tableCE1.RowCount);  //вычесляем результат для бокового
            result.Text = scientificResearch.DecisionResult().ToString(); //вычисляем результат общий
        }
    }
}
