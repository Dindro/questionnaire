using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace Project
{
    class BallEvent:Ball
    {
        public List<Event> events;
        public BallEvent(double ball)
        {
            this.ball = ball;
            events = new List<Event>();
        }

        public string DecisionResult(int RowCount)
        {
            result = ball * RowCount;
            if (result == 0)
                return "";
            else
                return result.ToString();
        }

        public string SetEventsInDataBase()
        {
            string randomId = String.Format("even_{0}_{1}", DateTime.Now.ToString("yyyyMMddHHmmssfff"), Guid.NewGuid());

            foreach (Event even in this.events)
            {
                OleDbCommand command = new OleDbCommand("INSERT INTO Event (Id,Название, Дата)VALUES(@Id, @Название, @Дата)", Data.OleDbConnection);
                command.Parameters.AddWithValue("Id", randomId);
                command.Parameters.AddWithValue("Название", even.nameOfEvent);
                command.Parameters.AddWithValue("Дата", even.dateOfEvent);
                command.ExecuteNonQuery();
            }
            return randomId;
        }

        public List<Event> Crushing(string line)
        {
            OleDbDataReader sqlReaderA = null;
            OleDbCommand command = new OleDbCommand("SELECT Название, Дата FROM Event WHERE  '" + line + "'=Id", Data.OleDbConnection); // * считываем все колонки
            sqlReaderA = command.ExecuteReader();
            while (sqlReaderA.Read())
            {
                events.Add(new Event(sqlReaderA["Название"].ToString(), sqlReaderA["Дата"].ToString()));
            }
            if (sqlReaderA != null)
                sqlReaderA.Close();
            DecisionResult(events.Count);

            return events;
        }

    }

    class Ball
    {
        public double ball;
        public double result;

        public int[] CrushLine(string line) //проблема с пробелом
        {
            string[] pieces = line.Split(',');
            int[] numbers = new int[pieces.Length-1];
            for (int i = 0; i < pieces.Length-1; i++)
            {
                numbers[i] = Convert.ToInt32(pieces[i]);
            }
            return numbers;
        }

        public void DeleteEvents(string [] lines, string nameOfDB)
        {
            for (int y = 0; y < lines.Length; y++)
            {
                OleDbCommand command = new OleDbCommand("DELETE FROM " + nameOfDB + " WHERE '" + lines[y] + "' =Id", Data.OleDbConnection);
                command.ExecuteNonQuery();
            }
        }
    }

    class Event
    {
        public string nameOfEvent;
        public string dateOfEvent;
        public Event(string nameOfEvent, string dateOfEvent)
        {
            this.nameOfEvent=nameOfEvent;
            this.dateOfEvent = dateOfEvent;
        }
    }    
}
