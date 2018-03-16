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
    public partial class СulturalСreativeActivity : UserControl
    {
        private static СulturalСreativeActivity instance;
        public static СulturalСreativeActivity Instance
        {
            get
            {
                if (instance == null)
                    instance = new СulturalСreativeActivity();
                return instance;
            }
        }

        CulturalСreative culturalCreative;

        DataGridView[] tables;
        DataGridView[] tablesA;
        DataGridView[] tablesB;
        List<Control> fields;

        public СulturalСreativeActivity()
        {
            InitializeComponent();
            Data.form.Text = "Культурно-творческая деятельность";
            panelMain.AutoScroll = true;
            culturalCreative = new CulturalСreative(new double[] { 2, 3, 4, 5, 6 },
                                                    new double[] { 2, 3, 4, 5, 6 },
                                                    new double[] {8});
            fields = new List<Control>() {};

            tables = new DataGridView[] {tableE1,tableD1,
                                           tableE2,tableD2,
                                           tableE3,tableD3,
                                           tableE4,tableD4,
                                           tableE5,tableD5,};

            tablesA = new DataGridView[] {tableAE1,tableAD1,
                                           tableAE2,tableAD2,
                                           tableAE3,tableAD3,
                                           tableAE4,tableAD4,
                                           tableAE5,tableAD5};

            tablesB = new DataGridView[] { tableBE1, tableBD1 };

            fields.AddRange(tables);
            fields.AddRange(tablesA);
            fields.AddRange(tablesB);
            result.Text = "0";
        }

        private new void Enter(object sender, EventArgs e)
        {
            Data.Enter((Control)sender);
        }

        #region Пункт #4.1
       
        #endregion

        #region Пункт #4.2
        private void plusB1_Click(object sender, EventArgs e)
        {
            culturalCreative.Plus(tableBE1, tableBD1);
            ballB1.Text = culturalCreative.participations[0].DecisionResult(tableBE1.RowCount);
            result.Text = culturalCreative.DecisionResult().ToString();
        }

        private void minusB1_Click(object sender, EventArgs e)
        {
            culturalCreative.Minus(tableBE1, tableBD1); //удаляем мероприятие из таблицы
            ballB1.Text = culturalCreative.participations[0].DecisionResult(tableBE1.RowCount);  //вычесляем результат для бокового
            result.Text = culturalCreative.DecisionResult().ToString(); //вычисляем результат общий
        }
        #endregion

        #region Пункт #4.3
        private void plus1_Click(object sender, EventArgs e)
        {
            culturalCreative.Plus(tableE1, tableD1);
            ball1.Text = culturalCreative.levels[0].DecisionResult(tableE1.RowCount);
            result.Text = culturalCreative.DecisionResult().ToString();
        }

        private void minus1_Click(object sender, EventArgs e)
        {
            culturalCreative.Minus(tableE1, tableD1); //удаляем мероприятие из таблицы
            ball1.Text = culturalCreative.levels[0].DecisionResult(tableE1.RowCount);  //вычесляем результат для бокового
            result.Text = culturalCreative.DecisionResult().ToString(); //вычисляем результат общий
        }

        private void plus2_Click(object sender, EventArgs e)
        {
            culturalCreative.Plus(tableE2, tableD2);
            ball2.Text = culturalCreative.levels[1].DecisionResult(tableE2.RowCount);
            result.Text = culturalCreative.DecisionResult().ToString();
        }

        private void minus2_Click(object sender, EventArgs e)
        {
            culturalCreative.Minus(tableE2, tableD2); //удаляем мероприятие из таблицы
            ball2.Text = culturalCreative.levels[1].DecisionResult(tableE2.RowCount);  //вычесляем результат для бокового
            result.Text = culturalCreative.DecisionResult().ToString(); //вычисляем результат общий
        }

        private void plus3_Click(object sender, EventArgs e)
        {
            culturalCreative.Plus(tableE3, tableD3);
            ball3.Text = culturalCreative.levels[2].DecisionResult(tableE3.RowCount);
            result.Text = culturalCreative.DecisionResult().ToString();
        }

        private void minus3_Click(object sender, EventArgs e)
        {
            culturalCreative.Minus(tableE3, tableD3); //удаляем мероприятие из таблицы
            ball3.Text = culturalCreative.levels[2].DecisionResult(tableE3.RowCount);  //вычесляем результат для бокового
            result.Text = culturalCreative.DecisionResult().ToString(); //вычисляем результат общий
        }

        private void plus4_Click(object sender, EventArgs e)
        {
            culturalCreative.Plus(tableE4, tableD4);
            ball4.Text = culturalCreative.levels[3].DecisionResult(tableE4.RowCount);
            result.Text = culturalCreative.DecisionResult().ToString();
        }

        private void minus4_Click(object sender, EventArgs e)
        {
            culturalCreative.Minus(tableE4, tableD4); //удаляем мероприятие из таблицы
            ball4.Text = culturalCreative.levels[3].DecisionResult(tableE4.RowCount);  //вычесляем результат для бокового
            result.Text = culturalCreative.DecisionResult().ToString(); //вычисляем результат общий
        }

        private void plus5_Click(object sender, EventArgs e)
        {
            culturalCreative.Plus(tableE5, tableD5);
            ball5.Text = culturalCreative.levels[4].DecisionResult(tableE5.RowCount);
            result.Text = culturalCreative.DecisionResult().ToString();
        }

        private void minus5_Click(object sender, EventArgs e)
        {
            culturalCreative.Minus(tableE5, tableD5); //удаляем мероприятие из таблицы
            ball5.Text = culturalCreative.levels[4].DecisionResult(tableE5.RowCount);  //вычесляем результат для бокового
            result.Text = culturalCreative.DecisionResult().ToString(); //вычисляем результат общий
        }

        #endregion

        #region Пункт #5.3
        private void plusA1_Click(object sender, EventArgs e)
        {
            culturalCreative.Plus(tableAE1, tableAD1);
            ballA1.Text = culturalCreative.performances[0].DecisionResult(tableAE1.RowCount);
            result.Text = culturalCreative.DecisionResult().ToString();
        }

        private void minusA1_Click(object sender, EventArgs e)
        {
            culturalCreative.Minus(tableAE1, tableAD1); //удаляем мероприятие из таблицы
            ballA1.Text = culturalCreative.performances[0].DecisionResult(tableAE1.RowCount);  //вычесляем результат для бокового
            result.Text = culturalCreative.DecisionResult().ToString(); //вычисляем результат общий
        }

        private void plusA2_Click(object sender, EventArgs e)
        {
            culturalCreative.Plus(tableAE2, tableAD2);
            ballA2.Text = culturalCreative.performances[1].DecisionResult(tableAE2.RowCount);
            result.Text = culturalCreative.DecisionResult().ToString();
        }

        private void minusA2_Click(object sender, EventArgs e)
        {
            culturalCreative.Minus(tableAE2, tableAD2); //удаляем мероприятие из таблицы
            ballA2.Text = culturalCreative.performances[1].DecisionResult(tableAE2.RowCount);  //вычесляем результат для бокового
            result.Text = culturalCreative.DecisionResult().ToString(); //вычисляем результат общий
        }

        private void plusA3_Click(object sender, EventArgs e)
        {
            culturalCreative.Plus(tableAE3, tableAD3);
            ballA3.Text = culturalCreative.performances[2].DecisionResult(tableAE3.RowCount);
            result.Text = culturalCreative.DecisionResult().ToString();
        }

        private void minusA3_Click(object sender, EventArgs e)
        {
            culturalCreative.Minus(tableAE3, tableAD3); //удаляем мероприятие из таблицы
            ballA3.Text = culturalCreative.performances[2].DecisionResult(tableAE3.RowCount);  //вычесляем результат для бокового
            result.Text = culturalCreative.DecisionResult().ToString(); //вычисляем результат общий
        }

        private void plusA4_Click(object sender, EventArgs e)
        {
            culturalCreative.Plus(tableAE4, tableAD4);
            ballA4.Text = culturalCreative.performances[3].DecisionResult(tableAE4.RowCount);
            result.Text = culturalCreative.DecisionResult().ToString();
        }

        private void minusA4_Click(object sender, EventArgs e)
        {
            culturalCreative.Minus(tableAE4, tableAD4); //удаляем мероприятие из таблицы
            ballA4.Text = culturalCreative.performances[3].DecisionResult(tableAE4.RowCount);  //вычесляем результат для бокового
            result.Text = culturalCreative.DecisionResult().ToString(); //вычисляем результат общий
        }

        private void plusA5_Click(object sender, EventArgs e)
        {
            culturalCreative.Plus(tableAE5, tableAD5);
            ballA5.Text = culturalCreative.performances[4].DecisionResult(tableAE5.RowCount);
            result.Text = culturalCreative.DecisionResult().ToString();
        }

        private void minusA5_Click(object sender, EventArgs e)
        {
            culturalCreative.Minus(tableAE5, tableAD5); //удаляем мероприятие из таблицы
            ballA5.Text = culturalCreative.performances[4].DecisionResult(tableAE5.RowCount);  //вычесляем результат для бокового
            result.Text = culturalCreative.DecisionResult().ToString(); //вычисляем результат общий
        }
        #endregion

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
            culturalCreative.levels = culturalCreative.SetFromTables(culturalCreative.levels, tables);
            culturalCreative.performances = culturalCreative.SetFromTables(culturalCreative.performances, tablesA);
            culturalCreative.participations = culturalCreative.SetFromTables(culturalCreative.participations, tablesB); 
            #endregion



            OleDbCommand command = new OleDbCommand("INSERT INTO Culture (Уровень1,Уровень2,Уровень3,Уровень4,Уровень5,УровеньА1, УровеньА2, УровеньА3, УровеньА4, УровеньА5,УровеньВ1,Результат)" +
                "VALUES(@Уровень1,@Уровень2,@Уровень3,@Уровень4,@Уровень5,@УровеньА1, @УровеньА2, @УровеньА3, @УровеньА4, @УровеньА5,@УровеньВ1,@Результат)", Data.OleDbConnection);
            command.Parameters.AddWithValue("Уровень1", culturalCreative.levels[0].SetEventsInDataBase());
            command.Parameters.AddWithValue("Уровень2", culturalCreative.levels[1].SetEventsInDataBase());
            command.Parameters.AddWithValue("Уровень3", culturalCreative.levels[2].SetEventsInDataBase());
            command.Parameters.AddWithValue("Уровень4", culturalCreative.levels[3].SetEventsInDataBase());
            command.Parameters.AddWithValue("Уровень5", culturalCreative.levels[4].SetEventsInDataBase());

            command.Parameters.AddWithValue("УровеньА1", culturalCreative.performances[0].SetEventsInDataBase());
            command.Parameters.AddWithValue("УровеньА2", culturalCreative.performances[1].SetEventsInDataBase());
            command.Parameters.AddWithValue("УровеньА3", culturalCreative.performances[2].SetEventsInDataBase());
            command.Parameters.AddWithValue("УровеньА4", culturalCreative.performances[3].SetEventsInDataBase());
            command.Parameters.AddWithValue("УровеньА5", culturalCreative.performances[4].SetEventsInDataBase());

            command.Parameters.AddWithValue("УровеньВ1", culturalCreative.participations[0].SetEventsInDataBase());

            command.Parameters.AddWithValue("Результат", culturalCreative.result.ToString());

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


            culturalCreative.human.SetInDataBase(id);

            culturalCreative.SetInWord();


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
