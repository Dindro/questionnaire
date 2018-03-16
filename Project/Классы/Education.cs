using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OleDb;
using System.IO;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;

namespace Project
{
    class Education : Activity
    {
        public List<BallEvent> confessions;
        public override int Status
        {
            get
            {
                return status;
            }

            set
            {
                status = value;
                if (status == 1) progress = 6;
                else
                    progress = 0;
            }
        }

        public Education(double[] levelBalls,
                         double[] levelBallsA)
        {
            confessions = new List<BallEvent>() { new BallEvent(levelBallsA[0]),
                                                         new BallEvent(levelBallsA[1]),
                                                         new BallEvent(levelBallsA[2]),
                                                         new BallEvent(levelBallsA[3]),
                                                         new BallEvent(levelBallsA[4]) };

            levels = new List<BallEvent>() { new BallEvent(levelBalls[0]),
                                             new BallEvent(levelBalls[1]),
                                             new BallEvent(levelBalls[2]),
                                             new BallEvent(levelBalls[3]),
                                             new BallEvent(levelBalls[4])};
        }

        public override double DecisionResult()
        {
            base.DecisionResult();
            foreach (BallEvent confession in confessions)
            {
                result += confession.result;
            }
            return result;
        }


        public void SetInWord()
        {
            File.WriteAllBytes("education.docx", Properties.Resources.education);
            Word.Application oWord = new Word.Application();
            Word.Document oDoc = oWord.Documents.Add(Environment.CurrentDirectory + "\\education.docx");
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

            oDoc.Bookmarks["PERIOD"].Range.Text = period;
            if (status == 1)
                oDoc.Bookmarks["S1"].Range.Text = "[x]";

            oDoc.Bookmarks["PROGRESS"].Range.Text = progress.ToString();
            oDoc.Bookmarks["RESULT"].Range.Text = result.ToString();
            LevelsSetInWord(oDoc, levels, bookMarks);
            LevelsSetInWord(oDoc, confessions, bookMarksA);
            Data.SaveFileDialog(oDoc);
            oWord.Quit();
        }

        public void SetInObjectFromDataBase(int id)
        {
            OleDbDataReader sqlReader = null;
            OleDbCommand command = new OleDbCommand("SELECT Education.Период as Период, Education.Статус as Статус, Education.Уровень1 as Уровень1, Education.Уровень2 as Уровень2, Education.Уровень3 as Уровень3, Education.Уровень4 as Уровень4, Education.Уровень5 as Уровень5, Education.УровеньА1 as УровеньА1, Education.УровеньА2 as УровеньА2, Education.УровеньА3 as УровеньА3, Education.УровеньА4 as УровеньА4, Education.УровеньА5 as УровеньА5, Education.Результат as Результат FROM Education, Human WHERE " + id.ToString() + "=Human.Id AND Human.IdActivity=Education.Id", Data.OleDbConnection); // * считываем все колонки
            sqlReader = command.ExecuteReader();
            string[] columnValues = new string[5];
            string[] columnValuesA = new string[5];
            while (sqlReader.Read())
            {
                period = sqlReader["Период"].ToString();
                Status = Convert.ToInt32(sqlReader["Статус"]);
                columnValues = new string[] { sqlReader["Уровень1"].ToString(), sqlReader["Уровень2"].ToString(), sqlReader["Уровень3"].ToString(), sqlReader["Уровень4"].ToString(), sqlReader["Уровень5"].ToString() };
                columnValuesA = new string[] { sqlReader["УровеньА1"].ToString(), sqlReader["УровеньА2"].ToString(), sqlReader["УровеньА3"].ToString(), sqlReader["УровеньА4"].ToString(), sqlReader["УровеньА5"].ToString() };
                result = Convert.ToDouble(sqlReader["Результат"]);
            }
            if (sqlReader != null)
                sqlReader.Close();

            levels[0].Crushing(columnValues[0]);
            levels[1].Crushing(columnValues[1]);
            levels[2].Crushing(columnValues[2]);
            levels[3].Crushing(columnValues[3]);
            levels[4].Crushing(columnValues[4]);

            confessions[0].Crushing(columnValuesA[0]);
            confessions[1].Crushing(columnValuesA[1]);
            confessions[2].Crushing(columnValuesA[2]);
            confessions[3].Crushing(columnValuesA[3]);
            confessions[4].Crushing(columnValuesA[4]);
        }

        public static void DeleteDataBase(int id)
        {
            OleDbDataReader sqlReader = null;
            OleDbCommand command = new OleDbCommand("SELECT Education.Id as IdEducation, Education.Уровень1 as Уровень1, Education.Уровень2 as Уровень2, Education.Уровень3 as Уровень3, Education.Уровень4 as Уровень4, Education.Уровень5 as Уровень5, Education.УровеньА1 as УровеньА1, Education.УровеньА2 as УровеньА2, Education.УровеньА3 as УровеньА3, Education.УровеньА4 as УровеньА4, Education.УровеньА5 as УровеньА5  FROM Education, Human WHERE " + id.ToString() + "=Human.Id AND Human.IdActivity=Education.Id", Data.OleDbConnection); // * считываем все колонки
            sqlReader = command.ExecuteReader();
            string[] columnValues = new string[5];
            string[] columnValuesA = new string[5];
            int idEducation = 0;
            while (sqlReader.Read())
            {
                idEducation = Convert.ToInt32(sqlReader["IdEducation"]);
                columnValues = new string[] { sqlReader["Уровень1"].ToString(), sqlReader["Уровень2"].ToString(), sqlReader["Уровень3"].ToString(), sqlReader["Уровень4"].ToString(), sqlReader["Уровень5"].ToString() };
                columnValuesA = new string[] { sqlReader["УровеньА1"].ToString(), sqlReader["УровеньА2"].ToString(), sqlReader["УровеньА3"].ToString(), sqlReader["УровеньА4"].ToString(), sqlReader["УровеньА5"].ToString() };
            }
            if (sqlReader != null)
                sqlReader.Close();

            Ball example = new Ball();
            example.DeleteEvents(columnValues, "Event");
            example.DeleteEvents(columnValuesA, "Event");

            command = new OleDbCommand("DELETE FROM Education WHERE " + idEducation + "=Id", Data.OleDbConnection);
            command.ExecuteNonQuery();
            command = new OleDbCommand("DELETE FROM Human WHERE " + id + "=Id", Data.OleDbConnection);
            command.ExecuteNonQuery();
        }
    }
}
