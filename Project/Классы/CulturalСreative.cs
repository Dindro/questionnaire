using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;
using System.IO;

namespace Project
{
    class CulturalСreative : Activity
    {
        public List<BallEvent> performances;
        public List<BallEvent> participations;
        public CulturalСreative(double[] levelBalls,
                                double[] levelBallsA,
                                double[] levelBallsB)
        {
            participations = new List<BallEvent>() { new BallEvent(levelBallsB[0])};
            levels = new List<BallEvent>() { new BallEvent(levelBalls[0]),
                                             new BallEvent(levelBalls[1]),
                                             new BallEvent(levelBalls[2]),
                                             new BallEvent(levelBalls[3]),
                                             new BallEvent(levelBalls[4])};
            performances = new List<BallEvent>() { new BallEvent(levelBallsA[0]),
                                             new BallEvent(levelBallsA[1]),
                                             new BallEvent(levelBallsA[2]),
                                             new BallEvent(levelBallsA[3]),
                                             new BallEvent(levelBallsA[4])};
        }

        public override double DecisionResult()
        {
            base.DecisionResult();
            foreach (BallEvent participation in participations)
            {
                result += participation.result;
            }
            foreach (BallEvent performance in performances)
            {
                result += performance.result;
            }
            return result;
        }

        public void SetInWord()
        {
            File.WriteAllBytes("culture.docx", Properties.Resources.culture);
            Word.Application oWord = new Word.Application();
            Word.Document oDoc = oWord.Documents.Add(Environment.CurrentDirectory + "\\culture.docx");

            string[,] bookMarks = new string[,] { { "E1", "D1", "B1" },
                                               { "E2", "D2", "B2" },
                                               { "E3", "D3", "B3" },
                                               { "E4", "D4", "B4" },
                                               { "E5", "D5", "B5" }};
            string[,] bookMarksA = new string[,] { { "AE1", "AD1", "AB1" },
                                               { "AE2", "AD2", "AB2" },
                                               { "AE3", "AD3", "AB3" },
                                               { "AE4", "AD4", "AB4" },
                                               { "AE5", "AD5", "AB5" }};
            string[,] bookMarksB = new string[,] { { "BE1", "BD1", "BB1" }};
            
            LevelsSetInWord(oDoc, levels, bookMarks);
            LevelsSetInWord(oDoc, performances, bookMarksA);
            LevelsSetInWord(oDoc, participations, bookMarksB);

            oDoc.Bookmarks["RESULT"].Range.Text = result.ToString();
            Data.SaveFileDialog(oDoc);
            oWord.Quit();
        }

        public void SetInObjectFromDataBase(int id)
        {
            OleDbDataReader sqlReader = null;
            OleDbCommand command = new OleDbCommand("SELECT Culture.Уровень1 as Уровень1, Culture.Уровень2 as Уровень2, Culture.Уровень3 as Уровень3, Culture.Уровень4 as Уровень4, Culture.Уровень5 as Уровень5, Culture.УровеньА1 as УровеньА1, Culture.УровеньА2 as УровеньА2, Culture.УровеньА3 as УровеньА3, Culture.УровеньА4 as УровеньА4, Culture.УровеньА5 as УровеньА5, Culture.УровеньВ1 as УровеньВ1, Culture.Результат as Результат FROM Culture, Human WHERE " + id.ToString() + "=Human.Id AND Human.IdActivity=Culture.Id", Data.OleDbConnection); // * считываем все колонки
            sqlReader = command.ExecuteReader();
            string[] columnValues = new string[5];
            string[] columnValuesA = new string[5];
            string[] columnValuesB = new string[1];
            while (sqlReader.Read())
            {
                columnValues = new string[] { sqlReader["Уровень1"].ToString(), sqlReader["Уровень2"].ToString(), sqlReader["Уровень3"].ToString(), sqlReader["Уровень4"].ToString(), sqlReader["Уровень5"].ToString() };
                columnValuesA = new string[] { sqlReader["УровеньА1"].ToString(), sqlReader["УровеньА2"].ToString(), sqlReader["УровеньА3"].ToString(), sqlReader["УровеньА4"].ToString(), sqlReader["УровеньА5"].ToString() };
                columnValuesB = new string[] { sqlReader["УровеньВ1"].ToString()};
                result = Convert.ToDouble(sqlReader["Результат"]);
            }
            if (sqlReader != null)
                sqlReader.Close();

            levels[0].Crushing(columnValues[0]);
            levels[1].Crushing(columnValues[1]);
            levels[2].Crushing(columnValues[2]);
            levels[3].Crushing(columnValues[3]);
            levels[4].Crushing(columnValues[4]);

            performances[0].Crushing(columnValuesA[0]);
            performances[1].Crushing(columnValuesA[1]);
            performances[2].Crushing(columnValuesA[2]);
            performances[3].Crushing(columnValuesA[3]);
            performances[4].Crushing(columnValuesA[4]);

            participations[0].Crushing(columnValuesB[0]);
        }

        //дальше надо и базу  и ворд менять

        public static void DeleteDataBase(int id)
        {
            OleDbDataReader sqlReader = null;
            OleDbCommand command = new OleDbCommand("SELECT Culture.Id as IdCulture, Culture.Уровень1 as Уровень1, Culture.Уровень2 as Уровень2, Culture.Уровень3 as Уровень3, Culture.Уровень4 as Уровень4, Culture.Уровень5 as Уровень5, Culture.УровеньА1 as УровеньА1, Culture.УровеньА2 as УровеньА2, Culture.УровеньА3 as УровеньА3, Culture.УровеньА4 as УровеньА4, Culture.УровеньА5 as УровеньА5, Culture.УровеньВ1 as УровеньВ1 FROM Culture, Human WHERE " + id.ToString() + "=Human.Id AND Human.IdActivity=Culture.Id", Data.OleDbConnection); // * считываем все колонки
            sqlReader = command.ExecuteReader();
            string[] columnValues = new string[5];
            string[] columnValuesA = new string[5];
            string[] columnValuesB = new string[1];
            int idCulture = 0;
            while (sqlReader.Read())
            {
                idCulture = Convert.ToInt32(sqlReader["IdCulture"]);
                columnValues = new string[] { sqlReader["Уровень1"].ToString(), sqlReader["Уровень2"].ToString(), sqlReader["Уровень3"].ToString(), sqlReader["Уровень4"].ToString(), sqlReader["Уровень5"].ToString() };
                columnValuesA = new string[] { sqlReader["УровеньА1"].ToString(), sqlReader["УровеньА2"].ToString(), sqlReader["УровеньА3"].ToString(), sqlReader["УровеньА4"].ToString(), sqlReader["УровеньА5"].ToString() };
                columnValuesB = new string[] { sqlReader["УровеньВ1"].ToString() };
            }
            if (sqlReader != null)
                sqlReader.Close();
            Ball example = new Ball();
            example.DeleteEvents(columnValues, "Event");
            example.DeleteEvents(columnValuesA, "Event");
            example.DeleteEvents(columnValuesB, "Event");

            command = new OleDbCommand("DELETE FROM Culture WHERE " + idCulture + "=Id", Data.OleDbConnection);
            command.ExecuteNonQuery();
            command = new OleDbCommand("DELETE FROM Human WHERE " + id + "=Id", Data.OleDbConnection);
            command.ExecuteNonQuery();
        }
    }
}
