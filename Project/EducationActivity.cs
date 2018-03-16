using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Data.OleDb;
using System.Text;
using System.Windows.Forms;
using System.Collections;
using Microsoft.Office.Interop.Word;

namespace Project
{
    public partial class EducationActivity : UserControl
    {
        private static EducationActivity instance;
        public static EducationActivity Instance
        {
            get
            {
                if (instance == null)
                    instance = new EducationActivity();
                return instance;
            }
        }

        Education educationActivity;

        public DataGridView [] tables;
        public DataGridView[] tablesA;
        public List<Control> fields; 

        public EducationActivity()
        {
            InitializeComponent();
            Data.form.Text = "Учебная деятельность";
            panelMain.AutoScroll = true;
            educationActivity = new Education(new double[] { 2, 3, 4, 5, 6 },
                                              new double[] { 2, 3, 4, 5, 6 });
            fields = new List<Control>() { maskedTextBoxPeriod,
                                           checkBox1};
            tables = new DataGridView[] { tableE1, tableD1,
                                                         tableE2, tableD2,
                                                         tableE3, tableD3,
                                                         tableE4, tableD4,
                                                         tableE5, tableD5};
            tablesA = new DataGridView[] {tableAE1,tableAD1,
                                                         tableAE2,tableAD2,
                                                         tableAE3,tableAD3,
                                                         tableAE4,tableAD4,
                                                         tableAE5,tableAD5};
            fields.AddRange(tables);
            fields.AddRange(tablesA);
            result.Text = "0";
        }

        private new void Enter(object sender, EventArgs e)
        {
            Data.Enter((Control)sender);
        }

        #region Пункт #1.1
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                educationActivity.Status = 1;
                ocenka1.Text = educationActivity.progress.ToString();
                result.Text = educationActivity.DecisionResult().ToString();
            }
            else
            {
                educationActivity.Status = 0;
                ocenka1.Text = "";
                result.Text = educationActivity.DecisionResult().ToString();
            }
                
        }
        #endregion

        #region Пункт #1.2
        private void plus1_Click(object sender, EventArgs e)
        {
            educationActivity.Plus(tableE1, tableD1);
            ball1.Text = educationActivity.levels[0].DecisionResult(tableE1.RowCount);
            result.Text = educationActivity.DecisionResult().ToString();
        }

        private void minus1_Click(object sender, EventArgs e)
        {
            educationActivity.Minus(tableE1, tableD1); //удаляем мероприятие из таблицы
            ball1.Text = educationActivity.levels[0].DecisionResult(tableE1.RowCount);  //вычесляем результат для бокового
            result.Text = educationActivity.DecisionResult().ToString(); //вычисляем результат общий
        }

        private void plus2_Click(object sender, EventArgs e)
        {
            educationActivity.Plus(tableE2, tableD2);
            ball2.Text = educationActivity.levels[1].DecisionResult(tableE2.RowCount);
            result.Text = educationActivity.DecisionResult().ToString();
        }

        private void minus2_Click(object sender, EventArgs e)
        {
            educationActivity.Minus(tableE2, tableD2); //удаляем мероприятие из таблицы
            ball2.Text = educationActivity.levels[1].DecisionResult(tableE2.RowCount);  //вычесляем результат для бокового
            result.Text = educationActivity.DecisionResult().ToString(); //вычисляем результат общий
        }

        private void plus3_Click(object sender, EventArgs e)
        {
            educationActivity.Plus(tableE3, tableD3);
            ball3.Text = educationActivity.levels[2].DecisionResult(tableE3.RowCount);
            result.Text = educationActivity.DecisionResult().ToString();
        }

        private void minus3_Click(object sender, EventArgs e)
        {
            educationActivity.Minus(tableE3, tableD3); //удаляем мероприятие из таблицы
            ball3.Text = educationActivity.levels[2].DecisionResult(tableE3.RowCount);  //вычесляем результат для бокового
            result.Text = educationActivity.DecisionResult().ToString(); //вычисляем результат общий
        }

        private void plus4_Click(object sender, EventArgs e)
        {
            educationActivity.Plus(tableE4, tableD4);
            ball4.Text = educationActivity.levels[3].DecisionResult(tableE4.RowCount);
            result.Text = educationActivity.DecisionResult().ToString();
        }

        private void minus4_Click(object sender, EventArgs e)
        {
            educationActivity.Minus(tableE4, tableD4); //удаляем мероприятие из таблицы
            ball4.Text = educationActivity.levels[3].DecisionResult(tableE4.RowCount);  //вычесляем результат для бокового
            result.Text = educationActivity.DecisionResult().ToString(); //вычисляем результат общий
        }

        private void plus5_Click(object sender, EventArgs e)
        {
            educationActivity.Plus(tableE5, tableD5);
            ball5.Text = educationActivity.levels[4].DecisionResult(tableE5.RowCount);
            result.Text = educationActivity.DecisionResult().ToString();
        }

        private void minus5_Click(object sender, EventArgs e)
        {
            educationActivity.Minus(tableE5, tableD5); //удаляем мероприятие из таблицы
            ball5.Text = educationActivity.levels[4].DecisionResult(tableE5.RowCount);  //вычесляем результат для бокового
            result.Text = educationActivity.DecisionResult().ToString(); //вычисляем результат общий
        }
        #endregion

        #region Пункт #1.3
        private void plusA1_Click(object sender, EventArgs e)
        {
            educationActivity.Plus(tableAE1, tableAD1); //добавляем в таблицу
            ballA1.Text = educationActivity.confessions[0].DecisionResult(tableAE1.RowCount);
            result.Text = educationActivity.DecisionResult().ToString();
        }

        private void minusA1_Click(object sender, EventArgs e)
        {
            educationActivity.Minus(tableAE1, tableAD1);
            ballA1.Text = educationActivity.confessions[0].DecisionResult(tableAE1.RowCount);
            result.Text = educationActivity.DecisionResult().ToString();
        }

        private void plusA2_Click(object sender, EventArgs e)
        {
            educationActivity.Plus(tableAE2, tableAD2); //добавляем в таблицу
            ballA2.Text = educationActivity.confessions[1].DecisionResult(tableAE2.RowCount);
            result.Text = educationActivity.DecisionResult().ToString();
        }

        private void minusA2_Click(object sender, EventArgs e)
        {
            educationActivity.Minus(tableAE2, tableAD2);
            ballA2.Text = educationActivity.confessions[1].DecisionResult(tableAE2.RowCount);
            result.Text = educationActivity.DecisionResult().ToString();
        }

        private void plusA3_Click(object sender, EventArgs e)
        {
            educationActivity.Plus(tableAE3, tableAD3); //добавляем в таблицу
            ballA3.Text = educationActivity.confessions[2].DecisionResult(tableAE3.RowCount);
            result.Text = educationActivity.DecisionResult().ToString();
        }

        private void minusA3_Click(object sender, EventArgs e)
        {
            educationActivity.Minus(tableAE3, tableAD3);
            ballA3.Text = educationActivity.confessions[2].DecisionResult(tableAE3.RowCount);
            result.Text = educationActivity.DecisionResult().ToString();
        }

        private void plusA4_Click(object sender, EventArgs e)
        {
            educationActivity.Plus(tableAE4, tableAD4); //добавляем в таблицу
            ballA4.Text = educationActivity.confessions[3].DecisionResult(tableAE4.RowCount);
            result.Text = educationActivity.DecisionResult().ToString();
        }

        private void minusA4_Click(object sender, EventArgs e)
        {
            educationActivity.Minus(tableAE4, tableAD4);
            ballA4.Text = educationActivity.confessions[3].DecisionResult(tableAE4.RowCount);
            result.Text = educationActivity.DecisionResult().ToString();
        }

        private void plusA5_Click(object sender, EventArgs e)
        {
            educationActivity.Plus(tableAE5, tableAD5); //добавляем в таблицу
            ballA5.Text = educationActivity.confessions[4].DecisionResult(tableAE5.RowCount);
            result.Text = educationActivity.DecisionResult().ToString();
        }

        private void minusA5_Click(object sender, EventArgs e)
        {
            educationActivity.Minus(tableAE5, tableAD5);
            ballA5.Text = educationActivity.confessions[4].DecisionResult(tableAE5.RowCount);
            result.Text = educationActivity.DecisionResult().ToString();
        }
        #endregion


        private void finish_Click(object sender, EventArgs e) //завершить
        {
            if (!Data.IsFilled(fields))
            {
                MessageBox.Show("Заполните пустые поля!", "Внимание");
                return;
            }

            #region Запись в объекты
            educationActivity.period = maskedTextBoxPeriod.Text;
            educationActivity.levels = educationActivity.SetFromTables(educationActivity.levels, tables);
            educationActivity.confessions = educationActivity.SetFromTables(educationActivity.confessions, tablesA);
            #endregion



            OleDbCommand command = new OleDbCommand("INSERT INTO Education (Период, Статус, Уровень1,Уровень2,Уровень3,Уровень4,Уровень5,УровеньА1, УровеньА2, УровеньА3, УровеньА4, УровеньА5, Результат)" +
                "VALUES(@Период, @Статус, @Уровень1,@Уровень2,@Уровень3,@Уровень4,@Уровень5,@УровеньА1, @УровеньА2, @УровеньА3, @УровеньА4, @УровеньА5, @Результат)", Data.OleDbConnection);
            command.Parameters.AddWithValue("Период", educationActivity.period.ToString());
            command.Parameters.AddWithValue("Статус", educationActivity.status.ToString());

            command.Parameters.AddWithValue("Уровень1", educationActivity.levels[0].SetEventsInDataBase());
            command.Parameters.AddWithValue("Уровень2", educationActivity.levels[1].SetEventsInDataBase());
            command.Parameters.AddWithValue("Уровень3", educationActivity.levels[2].SetEventsInDataBase());
            command.Parameters.AddWithValue("Уровень4", educationActivity.levels[3].SetEventsInDataBase());
            command.Parameters.AddWithValue("Уровень5", educationActivity.levels[4].SetEventsInDataBase());

            command.Parameters.AddWithValue("УровеньА1", educationActivity.confessions[0].SetEventsInDataBase());
            command.Parameters.AddWithValue("УровеньА2", educationActivity.confessions[1].SetEventsInDataBase());
            command.Parameters.AddWithValue("УровеньА3", educationActivity.confessions[2].SetEventsInDataBase());
            command.Parameters.AddWithValue("УровеньА4", educationActivity.confessions[3].SetEventsInDataBase());
            command.Parameters.AddWithValue("УровеньА5", educationActivity.confessions[4].SetEventsInDataBase());

            command.Parameters.AddWithValue("Результат", educationActivity.result.ToString());
            command.ExecuteNonQuery();


            int idEducation = 0;
            command = new OleDbCommand("SELECT @@IDENTITY AS id", Data.OleDbConnection);
            OleDbDataReader sqlReaderA = null;
            sqlReaderA = command.ExecuteReader();
            while (sqlReaderA.Read())
            {
                idEducation = Convert.ToInt32(sqlReaderA["id"]);
            }
            if (sqlReaderA != null)
                sqlReaderA.Close();

            educationActivity.human.SetInDataBase(idEducation); //запись человека в базу
            


            #region Запись в Word
            educationActivity.SetInWord();
            #endregion


            if (!Data.panel.Controls.Contains(Tables.Instance))
            {
                Data.panel.Controls.Add(Tables.Instance);
                Tables.Instance.Dock = DockStyle.Fill;
                Tables.Instance.BringToFront();
            }
            else
                Tables.Instance.BringToFront();
        }

        
    }
}
