using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Diagnostics;

namespace Project
{
    class Activity
    {
        public Human human;
        public double result;
        public List<BallEvent> levels;

        public string period;
        public double progress;
        public int status;
        public virtual int Status
        {
            get
            {
                return status;
            }
            set
            {
                status = value;
                if (status == 1) progress = 3;
                else if (status == 2) progress = 4;
                else progress = 6;
            }
        }

        public Activity()
        {
            human = Data.human;
            status = 0;
        }       

        public virtual double DecisionResult()
        {
            result = 0;
            result += progress;
            foreach (BallEvent level in levels)
            {
                result += level.result;
            }
            return result;
        }

        public virtual void Plus(DataGridView tableE, DataGridView tableD)
        {
            tableE.Rows.Add();
            tableD.Rows.Add();
            tableE[0, tableE.RowCount - 1].Value = tableE.RowCount;
            tableD[0, tableD.RowCount - 1].Value = tableD.RowCount;
        }

        public void Minus(DataGridView tableE, DataGridView tableD)
        {
            if (tableE.RowCount != 0)
            {
                tableE.Rows.RemoveAt(tableE.RowCount - 1);
                tableD.Rows.RemoveAt(tableD.RowCount - 1);
            }
        }

        public List<BallEvent> SetFromTables(List<BallEvent> levels, DataGridView[] tables)
        {
            int levelNumber = 0;
            for (int i = 0; i < tables.Length; i = i + 2)
            {
                for (int y = 0; y < tables[i].RowCount; y++)
                {
                    levels[levelNumber].events.Add(new Event(tables[i][1, y].Value.ToString(),
                                                             tables[i + 1][1, y].Value.ToString()));

                }
                levelNumber++;
            }
            return levels;
        }

        public void LevelsSetInWord(Document oDoc, List<BallEvent> levels, string[,] bookMarks)
        {
            int rowOfBookMarks = 0;
            foreach (BallEvent level in levels)
            {
                int count = 1;
                string name = "";
                string date = "";
                foreach (Event even in level.events)
                {
                    name += count + ") " + even.nameOfEvent + Environment.NewLine;
                    date += count + ") " + even.dateOfEvent + Environment.NewLine;
                    count++;
                }
                oDoc.Bookmarks[bookMarks[rowOfBookMarks, 0]].Range.Text = name;
                oDoc.Bookmarks[bookMarks[rowOfBookMarks, 1]].Range.Text = date;
                if (level.result != 0)
                    oDoc.Bookmarks[bookMarks[rowOfBookMarks, 2]].Range.Text = level.result.ToString();
                rowOfBookMarks++;
            }
        }

        public void GrantSetInWord(Document oDoc, Grant grant, string[] bookMarks)
        {
            if (grant.AreIs > 0)
            {
                oDoc.Bookmarks[bookMarks[0]].Range.Text = grant.name;
                oDoc.Bookmarks[bookMarks[1]].Range.Text = grant.date;
                oDoc.Bookmarks[bookMarks[2]].Range.Text = grant.progress.ToString();
            }
        }
    }
}
