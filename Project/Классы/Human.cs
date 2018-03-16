using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OleDb;
using System.Windows.Forms;
using System.IO;
using System.Collections;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;

namespace Project
{
    class Human
    {
        public string surname;
        public string name;
        public string fatherName;
        public string birthDay;
        public string university;
        public string specialty;
        public string preparation;
        public string programm;
        public int cours;
        public string phone;
        public string homePhone;
        public string email;
        public int activity;

        public void SetInWord()
        {
            File.WriteAllBytes("anketa.docx", Properties.Resources.anketa);
            Word.Application oWord = new Word.Application();
            Word.Document oDoc = oWord.Documents.Add(Environment.CurrentDirectory + "\\anketa.docx");
            oDoc.Bookmarks["SURNAME"].Range.Text = surname;
            oDoc.Bookmarks["NAME"].Range.Text = name;
            oDoc.Bookmarks["FATHERNAME"].Range.Text = fatherName;
            oDoc.Bookmarks["BIRTHDAY"].Range.Text = birthDay;
            oDoc.Bookmarks["UNIVERSITY"].Range.Text = university;
            oDoc.Bookmarks["SPECIALTY"].Range.Text = specialty;
            oDoc.Bookmarks["PREPARATION"].Range.Text = preparation;
            oDoc.Bookmarks["PROGRAMM"].Range.Text = programm;
            oDoc.Bookmarks["COURS"].Range.Text = cours.ToString();
            oDoc.Bookmarks["PHONE"].Range.Text = phone;
            oDoc.Bookmarks["HOMEPHONE"].Range.Text = homePhone;
            oDoc.Bookmarks["EMAIL"].Range.Text = email;
            if (activity == 1)
                oDoc.Bookmarks["D1"].Range.Text = "x";
            else if (activity == 2)
                oDoc.Bookmarks["D2"].Range.Text = "x";
            else if (activity == 3)
                oDoc.Bookmarks["D3"].Range.Text = "x";
            else if (activity == 4)
                oDoc.Bookmarks["D4"].Range.Text = "x";
            else
                oDoc.Bookmarks["D5"].Range.Text = "x";
            Data.SaveFileDialog(oDoc);
            oWord.Quit();
        }

        public void SetInObjectFromDataBase(int id)
        {
            OleDbDataReader sqlReader = null;
            OleDbCommand command = new OleDbCommand("SELECT Фамилия, Имя, Отчество, Рождение, Университет, Специальность, Подготовка, Программа,Курс, Телефон,Домашний,Почта,Деятельность FROM Human WHERE  " + id.ToString() + "=Id", Data.OleDbConnection); // * считываем все колонки
            sqlReader = command.ExecuteReader();

            while (sqlReader.Read())
            {
                this.surname = sqlReader["Фамилия"].ToString();
                this.name = sqlReader["Имя"].ToString();
                this.fatherName = sqlReader["Отчество"].ToString();
                this.birthDay = sqlReader["Рождение"].ToString();
                this.university = sqlReader["Университет"].ToString();
                this.specialty = sqlReader["Специальность"].ToString();
                this.preparation = sqlReader["Подготовка"].ToString();
                this.programm = sqlReader["Программа"].ToString();
                this.cours = Convert.ToInt32(sqlReader["Курс"]);
                this.phone = sqlReader["Телефон"].ToString();
                this.homePhone = sqlReader["Домашний"].ToString();
                this.email = sqlReader["Почта"].ToString();
                this.activity = Convert.ToInt32(sqlReader["Деятельность"]);
            }
            if (sqlReader != null)
                sqlReader.Close();
        }

        public void SetInDataBase(int idActivity)
        {
            OleDbCommand command = new OleDbCommand("INSERT INTO [Human] (IdActivity, Фамилия, Имя, Отчество, Рождение, Университет, Специальность," +
                "Подготовка, Программа, Курс, Телефон, Домашний, Почта, Деятельность)" +
                "VALUES(@IdActivity, @Фамилия, @Имя, @Отчество, @Рождение, @Университет, @Специальность, " +
                "@Подготовка, @Программа, @Курс, @Телефон, @Домашний, @Почта, @Деятельность)", Data.OleDbConnection);

            command.Parameters.AddWithValue("IdActivity", idActivity.ToString());
            command.Parameters.AddWithValue("Фамилия", surname);
            command.Parameters.AddWithValue("Имя", name);
            command.Parameters.AddWithValue("Отчество", fatherName);
            command.Parameters.AddWithValue("Рождение", birthDay);
            command.Parameters.AddWithValue("Университет", university);
            command.Parameters.AddWithValue("Специальность", specialty);
            command.Parameters.AddWithValue("Подготовка", preparation);
            command.Parameters.AddWithValue("Программа", programm);
            command.Parameters.AddWithValue("Курс", cours);
            command.Parameters.AddWithValue("Телефон", phone);
            command.Parameters.AddWithValue("Домашний", homePhone);
            command.Parameters.AddWithValue("Почта", email);
            command.Parameters.AddWithValue("Деятельность", activity);
            command.ExecuteNonQuery();
            command = null;
        }
    }
}
