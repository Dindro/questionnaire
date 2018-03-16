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
    public partial class PublicActivity : UserControl
    {
        private static PublicActivity instance;
        public static PublicActivity Instance
        {
            get
            {
                if (instance == null)
                    instance = new PublicActivity();
                return instance;
            }
        }

        Public publics;
        List<Control> fields;
        DataGridView[] tables;
        DataGridView[] tablesA;

        public PublicActivity()
        {
            InitializeComponent();
            Data.form.Text = "Общественная деятельность";
            panelMain.AutoScroll = true;
            publics = new Public(new double[] { 2, 3,4,5,6},
                                 new double[] { 2, 3, 4, 5, 6 });

            fields = new List<Control>();
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
            fields.AddRange(tables);
            fields.AddRange(tablesA);
            result.Text = "0";
        }

        #region Пункт #2.2
        private void plus1_Click(object sender, EventArgs e)
        {
            publics.Plus(tableE1, tableD1);
            ball1.Text = publics.levels[0].DecisionResult(tableE1.RowCount);
            result.Text = publics.DecisionResult().ToString();
        }

        private void minus1_Click(object sender, EventArgs e)
        {
            publics.Minus(tableE1, tableD1); //удаляем мероприятие из таблицы
            ball1.Text = publics.levels[0].DecisionResult(tableE1.RowCount);  //вычесляем результат для бокового
            result.Text = publics.DecisionResult().ToString(); //вычисляем результат общий
        }

        private void plus2_Click(object sender, EventArgs e)
        {
            publics.Plus(tableE2, tableD2);
            ball2.Text = publics.levels[1].DecisionResult(tableE2.RowCount);
            result.Text = publics.DecisionResult().ToString();
        }

        private void minus2_Click(object sender, EventArgs e)
        {
            publics.Minus(tableE2, tableD2); //удаляем мероприятие из таблицы
            ball2.Text = publics.levels[1].DecisionResult(tableE2.RowCount);  //вычесляем результат для бокового
            result.Text = publics.DecisionResult().ToString(); //вычисляем результат общий
        }

        private void plus3_Click(object sender, EventArgs e)
        {
            publics.Plus(tableE3, tableD3);
            ball3.Text = publics.levels[2].DecisionResult(tableE3.RowCount);
            result.Text = publics.DecisionResult().ToString();
        }

        private void minus3_Click(object sender, EventArgs e)
        {
            publics.Minus(tableE3, tableD3); //удаляем мероприятие из таблицы
            ball3.Text = publics.levels[2].DecisionResult(tableE3.RowCount);  //вычесляем результат для бокового
            result.Text = publics.DecisionResult().ToString(); //вычисляем результат общий
        }

        private void plus4_Click(object sender, EventArgs e)
        {
            publics.Plus(tableE4, tableD4);
            ball4.Text = publics.levels[3].DecisionResult(tableE4.RowCount);
            result.Text = publics.DecisionResult().ToString();
        }

        private void minus4_Click(object sender, EventArgs e)
        {
            publics.Minus(tableE4, tableD4); //удаляем мероприятие из таблицы
            ball4.Text = publics.levels[3].DecisionResult(tableE4.RowCount);  //вычесляем результат для бокового
            result.Text = publics.DecisionResult().ToString(); //вычисляем результат общий
        }

        private void plus5_Click(object sender, EventArgs e)
        {
            publics.Plus(tableE5, tableD5);
            ball5.Text = publics.levels[4].DecisionResult(tableE5.RowCount);
            result.Text = publics.DecisionResult().ToString();
        }

        private void minus5_Click(object sender, EventArgs e)
        {
            publics.Minus(tableE5, tableD5); //удаляем мероприятие из таблицы
            ball5.Text = publics.levels[4].DecisionResult(tableE5.RowCount);  //вычесляем результат для бокового
            result.Text = publics.DecisionResult().ToString(); //вычисляем результат общий
        }
        #endregion

        #region Пункт #5.3
        private void plusA1_Click(object sender, EventArgs e)
        {
            publics.Plus(tableAE1, tableAD1);
            ballA1.Text = publics.performances[0].DecisionResult(tableAE1.RowCount);
            result.Text = publics.DecisionResult().ToString();
        }

        private void minusA1_Click(object sender, EventArgs e)
        {
            publics.Minus(tableAE1, tableAD1); //удаляем мероприятие из таблицы
            ballA1.Text = publics.performances[0].DecisionResult(tableAE1.RowCount);  //вычесляем результат для бокового
            result.Text = publics.DecisionResult().ToString(); //вычисляем результат общий
        }

        private void plusA2_Click(object sender, EventArgs e)
        {
            publics.Plus(tableAE2, tableAD2);
            ballA2.Text = publics.performances[1].DecisionResult(tableAE2.RowCount);
            result.Text = publics.DecisionResult().ToString();
        }

        private void minusA2_Click(object sender, EventArgs e)
        {
            publics.Minus(tableAE2, tableAD2); //удаляем мероприятие из таблицы
            ballA2.Text = publics.performances[1].DecisionResult(tableAE2.RowCount);  //вычесляем результат для бокового
            result.Text = publics.DecisionResult().ToString(); //вычисляем результат общий
        }

        private void plusA3_Click(object sender, EventArgs e)
        {
            publics.Plus(tableAE3, tableAD3);
            ballA3.Text = publics.performances[2].DecisionResult(tableAE3.RowCount);
            result.Text = publics.DecisionResult().ToString();
        }

        private void minusA3_Click(object sender, EventArgs e)
        {
            publics.Minus(tableAE3, tableAD3); //удаляем мероприятие из таблицы
            ballA3.Text = publics.performances[2].DecisionResult(tableAE3.RowCount);  //вычесляем результат для бокового
            result.Text = publics.DecisionResult().ToString(); //вычисляем результат общий
        }

        private void plusA4_Click(object sender, EventArgs e)
        {
            publics.Plus(tableAE4, tableAD4);
            ballA4.Text = publics.performances[3].DecisionResult(tableAE4.RowCount);
            result.Text = publics.DecisionResult().ToString();
        }

        private void minusA4_Click(object sender, EventArgs e)
        {
            publics.Minus(tableAE4, tableAD4); //удаляем мероприятие из таблицы
            ballA4.Text = publics.performances[3].DecisionResult(tableAE4.RowCount);  //вычесляем результат для бокового
            result.Text = publics.DecisionResult().ToString(); //вычисляем результат общий
        }

        private void plusA5_Click(object sender, EventArgs e)
        {
            publics.Plus(tableAE5, tableAD5);
            ballA5.Text = publics.performances[4].DecisionResult(tableAE5.RowCount);
            result.Text = publics.DecisionResult().ToString();
        }

        private void minusA5_Click(object sender, EventArgs e)
        {
            publics.Minus(tableAE5, tableAD5); //удаляем мероприятие из таблицы
            ballA5.Text = publics.performances[4].DecisionResult(tableAE5.RowCount);  //вычесляем результат для бокового
            result.Text = publics.DecisionResult().ToString(); //вычисляем результат общий
        }
        #endregion

        private new void Enter(object sender, EventArgs e)
        {
            Data.Enter((Control)sender);
        }
        
        private void finish_Click(object sender, EventArgs e)
        {
            #region Проверка на заполненность
            if (!Data.IsFilled(fields))
            {
                MessageBox.Show("Заполните пустые поля!", "Внимание");
                return;
            }
            #endregion

            #region Запись в объекты
            publics.levels = publics.SetFromTables(publics.levels, tables);
            publics.performances = publics.SetFromTables(publics.performances, tablesA);
            #endregion

            OleDbCommand command = new OleDbCommand("INSERT INTO Publics (Уровень1,Уровень2,Уровень3,Уровень4,Уровень5, УровеньА1, УровеньА2, УровеньА3, УровеньА4, УровеньА5,Результат)" +
                "VALUES(@Уровень1,@Уровень2,@Уровень3,@Уровень4,@Уровень5, @УровеньА1, @УровеньА2, @УровеньА3, @УровеньА4, @УровеньА5,@Результат)", Data.OleDbConnection);

            command.Parameters.AddWithValue("Уровень1", publics.levels[0].SetEventsInDataBase());
            command.Parameters.AddWithValue("Уровень2", publics.levels[1].SetEventsInDataBase());
            command.Parameters.AddWithValue("Уровень3", publics.levels[2].SetEventsInDataBase());
            command.Parameters.AddWithValue("Уровень4", publics.levels[3].SetEventsInDataBase());
            command.Parameters.AddWithValue("Уровень5", publics.levels[4].SetEventsInDataBase());

            command.Parameters.AddWithValue("УровеньА1", publics.performances[0].SetEventsInDataBase());
            command.Parameters.AddWithValue("УровеньА2", publics.performances[1].SetEventsInDataBase());
            command.Parameters.AddWithValue("УровеньА3", publics.performances[2].SetEventsInDataBase());
            command.Parameters.AddWithValue("УровеньА4", publics.performances[3].SetEventsInDataBase());
            command.Parameters.AddWithValue("УровеньА5", publics.performances[4].SetEventsInDataBase());

            command.Parameters.AddWithValue("Результат", publics.result.ToString());
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

            publics.human.SetInDataBase(id);

            publics.SetInWord();

            if (!Data.panel.Controls.Contains(Tables.Instance)) //если нет в массиве то добавляем
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
