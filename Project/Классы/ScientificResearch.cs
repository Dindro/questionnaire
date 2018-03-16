using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OleDb;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Windows.Forms;

namespace Project
{
    class ScientificResearch : Activity
    {
        public List<BallEvent> performances;
        public List<BallEvent> documents;
        public List<BallEvent> grants;

        public ScientificResearch(double[] levelBalls,
                                  double[] levelBallsB,
                                  double[] levelBallsC,
                                  double[] levelBallsA)
        {
            levels = new List<BallEvent>() { new BallEvent(levelBalls[0]),
                                             new BallEvent(levelBalls[1]),
                                             new BallEvent(levelBalls[2]),
                                             new BallEvent(levelBalls[3]),
                                             new BallEvent(levelBalls[4])};
            documents = new List<BallEvent>() { new BallEvent(levelBallsB[0]) };
            grants = new List<BallEvent>() { new BallEvent(levelBallsC[0]) };
            performances = new List<BallEvent>() { new BallEvent(levelBallsA[0]),
                                             new BallEvent(levelBallsA[1]),
                                             new BallEvent(levelBallsA[2]),
                                             new BallEvent(levelBallsA[3]),
                                             new BallEvent(levelBallsA[4])};


        }

        public override double DecisionResult()
        {
            base.DecisionResult();
            foreach (BallEvent performance in performances)
            {
                result += performance.result;
            }
            foreach (BallEvent document in documents)
            {
                result += document.result;
            }
            foreach (BallEvent grant in grants)
            {
                result += grant.result;
            }
            return result;
        }

        public void SetInWord()
        {
            File.WriteAllBytes("scientific.docx", Properties.Resources.scientific);
            Word.Application oWord = new Word.Application();
            Word.Document oDoc = oWord.Documents.Add(Environment.CurrentDirectory + "\\scientific.docx");
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
            string[,] bookMarksB = new string[,] { { "DOCE", "DOCD", "DOCB" } };
            string[,] bookMarksC = new string[,] { { "GRANTE", "GRANTD", "GRANTB" } };
            oDoc.Bookmarks["PERIOD"].Range.Text = period;
            if (status == 1)
                oDoc.Bookmarks["S1"].Range.Text = "[x]";
            else if (status == 2)
                oDoc.Bookmarks["S2"].Range.Text = "[x]";
            else
                oDoc.Bookmarks["S3"].Range.Text = "[x]";
            oDoc.Bookmarks["PROGRESS"].Range.Text = progress.ToString();
            LevelsSetInWord(oDoc, levels, bookMarks);

            LevelsSetInWord(oDoc, documents, bookMarksB);
            LevelsSetInWord(oDoc, grants, bookMarksC);

            LevelsSetInWord(oDoc, performances, bookMarksA);

            oDoc.Bookmarks["RESULT"].Range.Text = result.ToString();

            Data.SaveFileDialog(oDoc);
            oWord.Quit();
        }

        public void SetInObjectFromDataBase(int id)
        {
            OleDbDataReader sqlReader = null;
            OleDbCommand command = new OleDbCommand("SELECT Scientific.Период as Период, Scientific.Статус as Статус, Scientific.Уровень1 as Уровень1, Scientific.Уровень2 as Уровень2, Scientific.Уровень3 as Уровень3, Scientific.Уровень4 as Уровень4, Scientific.Уровень5 as Уровень5, Scientific.УровеньВ1 as УровеньВ1, Scientific.УровеньС1 as УровеньС1, Scientific.УровеньА1 as УровеньА1, Scientific.УровеньА2 as УровеньА2, Scientific.УровеньА3 as УровеньА3, Scientific.УровеньА4 as УровеньА4, Scientific.УровеньА5 as УровеньА5, Scientific.Результат as Результат FROM Scientific, Human WHERE " + id.ToString() + "=Human.Id AND Human.IdActivity=Scientific.Id", Data.OleDbConnection); // * считываем все колонки
            sqlReader = command.ExecuteReader();
            string[] columnValues = new string[5];
            string[] columnValuesA = new string[5];
            string[] columnValuesB = new string[1];
            string[] columnValuesC = new string[1];
            while (sqlReader.Read())
            {
                period = sqlReader["Период"].ToString();
                Status = Convert.ToInt32(sqlReader["Статус"]);
                
                columnValues = new string[] { sqlReader["Уровень1"].ToString(), sqlReader["Уровень2"].ToString(), sqlReader["Уровень3"].ToString(), sqlReader["Уровень4"].ToString(), sqlReader["Уровень5"].ToString() };

                columnValuesB = new string[] { sqlReader["УровеньВ1"].ToString() };
                columnValuesC = new string[] { sqlReader["УровеньС1"].ToString() };

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

            documents[0].Crushing(columnValuesB[0]);
            grants[0].Crushing(columnValuesC[0]);

            performances[0].Crushing(columnValuesA[0]);
            performances[1].Crushing(columnValuesA[1]);
            performances[2].Crushing(columnValuesA[2]);
            performances[3].Crushing(columnValuesA[3]);
            performances[4].Crushing(columnValuesA[4]);
        }



        public static void DeleteDataBase(int id)
        {
            OleDbDataReader sqlReader = null;
            OleDbCommand command = new OleDbCommand("SELECT Scientific.Id as IdScientific, Scientific.Уровень1 as Уровень1, Scientific.Уровень2 as Уровень2, Scientific.Уровень3 as Уровень3, Scientific.Уровень4 as Уровень4, Scientific.Уровень5 as Уровень5, Scientific.УровеньВ1 as УровеньВ1, Scientific.УровеньС1 as УровеньС1, Scientific.УровеньА1 as УровеньА1, Scientific.УровеньА2 as УровеньА2, Scientific.УровеньА3 as УровеньА3, Scientific.УровеньА4 as УровеньА4, Scientific.УровеньА5 as УровеньА5 FROM Scientific, Human WHERE " + id.ToString() + "=Human.Id AND Human.IdActivity=Scientific.Id", Data.OleDbConnection); // * считываем все колонки
            sqlReader = command.ExecuteReader();
            string[] columnValues = new string[5];
            string[] columnValuesA = new string[5];
            string[] columnValuesB = new string[1];
            string[] columnValuesC = new string[1];
            int idScientific = 0;
            while (sqlReader.Read())
            {
                idScientific = Convert.ToInt32(sqlReader["IdScientific"]);
                columnValues = new string[] { sqlReader["Уровень1"].ToString(), sqlReader["Уровень2"].ToString(), sqlReader["Уровень3"].ToString(), sqlReader["Уровень4"].ToString(), sqlReader["Уровень5"].ToString() };
                columnValuesB = new string[] { sqlReader["УровеньВ1"].ToString() };
                columnValuesC = new string[] { sqlReader["УровеньС1"].ToString() };
                columnValuesA = new string[] { sqlReader["УровеньА1"].ToString(), sqlReader["УровеньА2"].ToString(), sqlReader["УровеньА3"].ToString(), sqlReader["УровеньА4"].ToString(), sqlReader["УровеньА5"].ToString() };
            }
            if (sqlReader != null)
                sqlReader.Close();
            Ball example = new Ball();
            example.DeleteEvents(columnValues, "Event");
            example.DeleteEvents(columnValuesA, "Event");
            example.DeleteEvents(columnValuesB, "Event");
            example.DeleteEvents(columnValuesC, "Event");

            command = new OleDbCommand("DELETE FROM Scientific WHERE " + idScientific + "=Id", Data.OleDbConnection);
            command.ExecuteNonQuery();
            command = new OleDbCommand("DELETE FROM Human WHERE " + id + "=Id", Data.OleDbConnection);
            command.ExecuteNonQuery();
        }
    }
}
