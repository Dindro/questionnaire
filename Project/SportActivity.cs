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
    public partial class SportActivity : UserControl
    {
        private static SportActivity instance;
        public static SportActivity Instance
        {
            get
            {
                if (instance == null)
                    instance = new SportActivity();
                return instance;
            }
        }

        Sport sport;

        List<Control> fields;
        DataGridView[] tables;
        DataGridView[] tablesA;


        public SportActivity()
        {
            InitializeComponent();
            Data.form.Text = "Спортивная деятельность";
            panelMain.AutoScroll = true;
            sport = new Sport(new double[] { 2, 3, 4, 5, 6 },
                              new double[] { 2, 3, 4, 5, 6 },
                              6);

            fields = new List<Control>() {};

            tables = new DataGridView[] {tableE1,tableD1,
                                           tableE2,tableD2,
                                           tableE3,tableD3,
                                           tableE4,tableD4,
                                           tableE5,tableD5 };

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

        #region Пункт #5.2
        private void plus1_Click(object sender, EventArgs e)
        {
            sport.Plus(tableE1, tableD1);
            ball1.Text = sport.levels[0].DecisionResult(tableE1.RowCount);
            result.Text = sport.DecisionResult().ToString();
        }

        private void minus1_Click(object sender, EventArgs e)
        {
            sport.Minus(tableE1, tableD1); //удаляем мероприятие из таблицы
            ball1.Text = sport.levels[0].DecisionResult(tableE1.RowCount);  //вычесляем результат для бокового
            result.Text = sport.DecisionResult().ToString(); //вычисляем результат общий
        }

        private void plus2_Click(object sender, EventArgs e)
        {
            sport.Plus(tableE2, tableD2);
            ball2.Text = sport.levels[1].DecisionResult(tableE2.RowCount);
            result.Text = sport.DecisionResult().ToString();
        }

        private void minus2_Click(object sender, EventArgs e)
        {
            sport.Minus(tableE2, tableD2); //удаляем мероприятие из таблицы
            ball2.Text = sport.levels[1].DecisionResult(tableE2.RowCount);  //вычесляем результат для бокового
            result.Text = sport.DecisionResult().ToString(); //вычисляем результат общий
        }

        private void plus3_Click(object sender, EventArgs e)
        {
            sport.Plus(tableE3, tableD3);
            ball3.Text = sport.levels[2].DecisionResult(tableE3.RowCount);
            result.Text = sport.DecisionResult().ToString();
        }

        private void minus3_Click(object sender, EventArgs e)
        {
            sport.Minus(tableE3, tableD3); //удаляем мероприятие из таблицы
            ball3.Text = sport.levels[2].DecisionResult(tableE3.RowCount);  //вычесляем результат для бокового
            result.Text = sport.DecisionResult().ToString(); //вычисляем результат общий
        }

        private void plus4_Click(object sender, EventArgs e)
        {
            sport.Plus(tableE4, tableD4);
            ball4.Text = sport.levels[3].DecisionResult(tableE4.RowCount);
            result.Text = sport.DecisionResult().ToString();
        }

        private void minus4_Click(object sender, EventArgs e)
        {
            sport.Minus(tableE4, tableD4); //удаляем мероприятие из таблицы
            ball4.Text = sport.levels[3].DecisionResult(tableE4.RowCount);  //вычесляем результат для бокового
            result.Text = sport.DecisionResult().ToString(); //вычисляем результат общий
        }

        private void plus5_Click(object sender, EventArgs e)
        {
            sport.Plus(tableE5, tableD5);
            ball5.Text = sport.levels[4].DecisionResult(tableE5.RowCount);
            result.Text = sport.DecisionResult().ToString();
        }

        private void minus5_Click(object sender, EventArgs e)
        {
            sport.Minus(tableE5, tableD5); //удаляем мероприятие из таблицы
            ball5.Text = sport.levels[4].DecisionResult(tableE5.RowCount);  //вычесляем результат для бокового
            result.Text = sport.DecisionResult().ToString(); //вычисляем результат общий
        }


        #endregion

        #region Пункт #5.3
        private void plusA1_Click(object sender, EventArgs e)
        {
            sport.Plus(tableAE1, tableAD1);
            ballA1.Text = sport.participations[0].DecisionResult(tableAE1.RowCount);
            result.Text = sport.DecisionResult().ToString();
        }

        private void minusA1_Click(object sender, EventArgs e)
        {
            sport.Minus(tableAE1, tableAD1); //удаляем мероприятие из таблицы
            ballA1.Text = sport.participations[0].DecisionResult(tableAE1.RowCount);  //вычесляем результат для бокового
            result.Text = sport.DecisionResult().ToString(); //вычисляем результат общий
        }

        private void plusA2_Click(object sender, EventArgs e)
        {
            sport.Plus(tableAE2, tableAD2);
            ballA2.Text = sport.participations[1].DecisionResult(tableAE2.RowCount);
            result.Text = sport.DecisionResult().ToString();
        }

        private void minusA2_Click(object sender, EventArgs e)
        {
            sport.Minus(tableAE2, tableAD2); //удаляем мероприятие из таблицы
            ballA2.Text = sport.participations[1].DecisionResult(tableAE2.RowCount);  //вычесляем результат для бокового
            result.Text = sport.DecisionResult().ToString(); //вычисляем результат общий
        }

        private void plusA3_Click(object sender, EventArgs e)
        {
            sport.Plus(tableAE3, tableAD3);
            ballA3.Text = sport.participations[2].DecisionResult(tableAE3.RowCount);
            result.Text = sport.DecisionResult().ToString();
        }

        private void minusA3_Click(object sender, EventArgs e)
        {
            sport.Minus(tableAE3, tableAD3); //удаляем мероприятие из таблицы
            ballA3.Text = sport.participations[2].DecisionResult(tableAE3.RowCount);  //вычесляем результат для бокового
            result.Text = sport.DecisionResult().ToString(); //вычисляем результат общий
        }

        private void plusA4_Click(object sender, EventArgs e)
        {
            sport.Plus(tableAE4, tableAD4);
            ballA4.Text = sport.participations[3].DecisionResult(tableAE4.RowCount);
            result.Text = sport.DecisionResult().ToString();
        }

        private void minusA4_Click(object sender, EventArgs e)
        {
            sport.Minus(tableAE4, tableAD4); //удаляем мероприятие из таблицы
            ballA4.Text = sport.participations[3].DecisionResult(tableAE4.RowCount);  //вычесляем результат для бокового
            result.Text = sport.DecisionResult().ToString(); //вычисляем результат общий
        }

        private void plusA5_Click(object sender, EventArgs e)
        {
            sport.Plus(tableAE5, tableAD5);
            ballA5.Text = sport.participations[4].DecisionResult(tableAE5.RowCount);
            result.Text = sport.DecisionResult().ToString();
        }

        private void minusA5_Click(object sender, EventArgs e)
        {
            sport.Minus(tableAE5, tableAD5); //удаляем мероприятие из таблицы
            ballA5.Text = sport.participations[4].DecisionResult(tableAE5.RowCount);  //вычесляем результат для бокового
            result.Text = sport.DecisionResult().ToString(); //вычисляем результат общий
        }
        #endregion

        private void finish_Click(object sender, EventArgs e)
        {
            #region Проверка на заполненность
            List<Control> controls = new List<Control>();
            if (gtoCheck.Checked)
            {
                controls.Add(gtoName);
                controls.Add(gtoDate);
            }
            controls.AddRange(fields);
            if (!Data.IsFilled(controls))
            {
                MessageBox.Show("Заполните пустые поля!", "Внимание");
                return;
            }
            #endregion

            #region Запись в объекты
            sport.levels = sport.SetFromTables(sport.levels, tables);
            sport.participations = sport.SetFromTables(sport.participations, tablesA);
            if (gtoCheck.Checked)
            {
                sport.gto.name = gtoName.Text;
                sport.gto.date = gtoDate.Text;
            }
            #endregion

            OleDbCommand command = new OleDbCommand("INSERT INTO Sports (Уровень1,Уровень2,Уровень3,Уровень4,Уровень5, УровеньА1, УровеньА2, УровеньА3, УровеньА4, УровеньА5, Гто, ГтоНазвание, ГтоДата, Результат)" +
               "VALUES(@Уровень1,@Уровень2,@Уровень3,@Уровень4,@Уровень5, @УровеньА1, @УровеньА2, @УровеньА3, @УровеньА4, @УровеньА5,@Гто, @ГтоНазвание, @ГтоДата, @Результат)", Data.OleDbConnection);

            command.Parameters.AddWithValue("Уровень1", sport.levels[0].SetEventsInDataBase());
            command.Parameters.AddWithValue("Уровень2", sport.levels[1].SetEventsInDataBase());
            command.Parameters.AddWithValue("Уровень3", sport.levels[2].SetEventsInDataBase());
            command.Parameters.AddWithValue("Уровень4", sport.levels[3].SetEventsInDataBase());
            command.Parameters.AddWithValue("Уровень5", sport.levels[4].SetEventsInDataBase());

            command.Parameters.AddWithValue("УровеньА1", sport.participations[0].SetEventsInDataBase());
            command.Parameters.AddWithValue("УровеньА2", sport.participations[1].SetEventsInDataBase());
            command.Parameters.AddWithValue("УровеньА3", sport.participations[2].SetEventsInDataBase());
            command.Parameters.AddWithValue("УровеньА4", sport.participations[3].SetEventsInDataBase());
            command.Parameters.AddWithValue("УровеньА5", sport.participations[4].SetEventsInDataBase());

            command.Parameters.AddWithValue("Гто", sport.gto.AreIs);
            command.Parameters.AddWithValue("ГтоНазвание", sport.gto.name);
            command.Parameters.AddWithValue("ГтоДата", sport.gto.date);

            command.Parameters.AddWithValue("Результат", sport.result.ToString());

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

            sport.human.SetInDataBase(id);

            sport.SetInWord();

            if (!Data.panel.Controls.Contains(Tables.Instance)) //если нет в массиве то добавляем
            {
                Data.panel.Controls.Add(Tables.Instance);
                Tables.Instance.Dock = DockStyle.Fill;
                Tables.Instance.BringToFront();
            }
            else
                Tables.Instance.BringToFront();

        }

        private void gtoCheck_CheckedChanged(object sender, EventArgs e)
        {
            if (gtoCheck.Checked)
            {
                gtoName.Enabled = true;
                gtoDate.Enabled = true;
                sport.gto.AreIs = 1;
                gtoBall.Text = sport.gto.progress.ToString();
            }
            else
            {
                gtoName.Enabled = false;
                gtoDate.Enabled = false;
                sport.gto.AreIs = 0;
                gtoBall.Text = "";
            }
            gtoName.BackColor = Color.Empty;
            gtoDate.BackColor = Color.Empty;
            result.Text = sport.DecisionResult().ToString();
        }
    }
}
