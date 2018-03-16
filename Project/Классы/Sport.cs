using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OleDb;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;
using System.IO;

namespace Project
{
    class Sport : Activity
    {
        public List<BallEvent> participations;
        public Grant gto;

        public Sport(double[] levelBalls,
                    double[] levelBallsA,
                    double gtoBall)
        {
            levels = new List<BallEvent>() { new BallEvent(levelBalls[0]),
                                             new BallEvent(levelBalls[1]),
                                             new BallEvent(levelBalls[2]),
                                             new BallEvent(levelBalls[3]),
                                             new BallEvent(levelBalls[4])};
            participations = new List<BallEvent>() { new BallEvent(levelBallsA[0]),
                                                   new BallEvent(levelBallsA[1]),
                                                   new BallEvent(levelBallsA[2]),
                                                   new BallEvent(levelBallsA[3]),
                                                   new BallEvent(levelBallsA[4])};
            gto = new Grant(gtoBall);
        }

        public override double DecisionResult()
        {
            base.DecisionResult();
            foreach (BallEvent participation in participations)
            {
                result += participation.result;
            }
            result += gto.progress;
            return result;
        }

        public void SetInWord()
        {
            File.WriteAllBytes("sport.docx", Properties.Resources.sport);
            Word.Application oWord = new Word.Application();
            Word.Document oDoc = oWord.Documents.Add(Environment.CurrentDirectory + "\\sport.docx");

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
            string[] bookMarksB = new string[] {"GTOE", "GTOD", "GTOB"};
            LevelsSetInWord(oDoc, levels, bookMarks);
            LevelsSetInWord(oDoc, participations, bookMarksA);
            GrantSetInWord(oDoc, gto, bookMarksB);
            oDoc.Bookmarks["RESULT"].Range.Text = result.ToString();

            Data.SaveFileDialog(oDoc);
            oWord.Quit();
        }

        public void SetInObjectFromDataBase(int id)
        {
            OleDbDataReader sqlReader = null;
            OleDbCommand command = new OleDbCommand("SELECT Sports.Уровень1 as Уровень1, Sports.Уровень2 as Уровень2, Sports.Уровень3 as Уровень3, Sports.Уровень4 as Уровень4, Sports.Уровень5 as Уровень5, Sports.УровеньА1 as УровеньА1, Sports.УровеньА2 as УровеньА2, Sports.УровеньА3 as УровеньА3, Sports.УровеньА4 as УровеньА4, Sports.УровеньА5 as УровеньА5, Sports.Гто as Гто, Sports.ГтоНазвание as ГтоНазвание, Sports.ГтоДата as ГтоДата, Sports.Результат as Результат  FROM Sports, Human WHERE " + id.ToString() + "=Human.Id AND Human.IdActivity=Sports.Id", Data.OleDbConnection); // * считываем все колонки
            sqlReader = command.ExecuteReader();
            string[] columnValues = new string[5];
            string[] columnValuesA = new string[5];
            while (sqlReader.Read())
            {
                columnValues = new string[] { sqlReader["Уровень1"].ToString(), sqlReader["Уровень2"].ToString(), sqlReader["Уровень3"].ToString(), sqlReader["Уровень4"].ToString(), sqlReader["Уровень5"].ToString() };
                columnValuesA = new string[] { sqlReader["УровеньА1"].ToString(), sqlReader["УровеньА2"].ToString(), sqlReader["УровеньА3"].ToString(), sqlReader["УровеньА4"].ToString(), sqlReader["УровеньА5"].ToString() };

                gto.AreIs = Convert.ToDouble(sqlReader["Гто"]);
                gto.name = sqlReader["ГтоНазвание"].ToString();
                gto.date = sqlReader["ГтоДата"].ToString();

                result = Convert.ToDouble(sqlReader["Результат"]);
            }
            if (sqlReader != null)
                sqlReader.Close();

            levels[0].Crushing(columnValues[0]);
            levels[1].Crushing(columnValues[1]);
            levels[2].Crushing(columnValues[2]);
            levels[3].Crushing(columnValues[3]);
            levels[4].Crushing(columnValues[4]);

            participations[0].Crushing(columnValuesA[0]);
            participations[1].Crushing(columnValuesA[1]);
            participations[2].Crushing(columnValuesA[2]);
            participations[3].Crushing(columnValuesA[3]);
            participations[4].Crushing(columnValuesA[4]);
        }

        public static void DeleteDataBase(int id)
        {
            OleDbDataReader sqlReader = null;
            OleDbCommand command = new OleDbCommand("SELECT Sports.Id as IdSports, Sports.Уровень1 as Уровень1, Sports.Уровень2 as Уровень2, Sports.Уровень3 as Уровень3, Sports.Уровень4 as Уровень4, Sports.Уровень5 as Уровень5, Sports.УровеньА1 as УровеньА1, Sports.УровеньА2 as УровеньА2, Sports.УровеньА3 as УровеньА3, Sports.УровеньА4 as УровеньА4, Sports.УровеньА5 as УровеньА5 FROM Sports, Human WHERE " + id.ToString() + "=Human.Id AND Human.IdActivity=Sports.Id", Data.OleDbConnection); // * считываем все колонки
            sqlReader = command.ExecuteReader();
            string[] columnValues = new string[5];
            string[] columnValuesA = new string[5];
            int idSports = 0;
            while (sqlReader.Read())
            {
                idSports = Convert.ToInt32(sqlReader["IdSports"]);
                columnValues = new string[] { sqlReader["Уровень1"].ToString(), sqlReader["Уровень2"].ToString(), sqlReader["Уровень3"].ToString(), sqlReader["Уровень4"].ToString(), sqlReader["Уровень5"].ToString() };
                columnValuesA = new string[] { sqlReader["УровеньА1"].ToString(), sqlReader["УровеньА2"].ToString(), sqlReader["УровеньА3"].ToString(), sqlReader["УровеньА4"].ToString(), sqlReader["УровеньА5"].ToString() };
            }
            if (sqlReader != null)
                sqlReader.Close();
            Ball example = new Ball();
            example.DeleteEvents(columnValues, "Event");
            example.DeleteEvents(columnValuesA, "Event");

            command = new OleDbCommand("DELETE FROM Sports WHERE " + idSports + "=Id", Data.OleDbConnection);
            command.ExecuteNonQuery();
            command = new OleDbCommand("DELETE FROM Human WHERE " + id + "=Id", Data.OleDbConnection);
            command.ExecuteNonQuery();
        }
    }
}
