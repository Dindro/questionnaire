using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;
using System.Data.OleDb;

namespace Project
{
    class Grant
    {
        public string name;
        public string date;
        public double progress;
        private double ball;
        private double areIs;
        public double AreIs
        {
            get { return areIs; }
            set
            {
                areIs = value;
                if (areIs > 0)
                    progress = ball;
                else
                    progress = 0;
            }
        }

        public Grant(double ball)
        {
            this.ball = ball;
            areIs = 0;
            name = "";
            date = "";
        }
    }

    class Login
    {
        private string password = "iadmin";
        public bool Check(string password)
        {
            if (this.password == password)
            {
                if (!Data.panel.Controls.Contains(Tables.Instance)) //если нет в массиве то добавляем
                {
                    Data.panel.Controls.Add(Tables.Instance);
                    Tables.Instance.Dock = DockStyle.Fill;
                    Tables.Instance.BringToFront();
                }
                else
                    Tables.Instance.BringToFront();
                Tables.Instance.delete.Visible = true;
                Tables.Instance.delete1.Visible = true;
                Tables.Instance.delete3.Visible = true;
                Tables.Instance.delete4.Visible = true;
                Tables.Instance.delete5.Visible = true;
                return true;
            }
            else return false;
        }
    }

   

    class Data
    {
        public static Form form;
        public static Panel panel;
        public static OleDbConnection OleDbConnection;
        public static Human human = new Human();

        public static bool IsProbel(string text)
        {
            int count = 0;
            for (int i = 0; i < text.Length; i++)
                if (text[i] == ' ')
                    count++;

            if (text.Length == count)
                return true;
            else
                return false;
        }

        public static void SaveFileDialog(Document oDoc)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Word (*.docx)|*.docx";
            DialogResult dialogResult = saveFileDialog1.ShowDialog();
            if (dialogResult == DialogResult.OK)
            {
                oDoc.SaveAs(FileName: saveFileDialog1.FileName);
                oDoc.Close();
            }
            else oDoc = null;
        }

        public static void SetInObjectFromDataBaseForTable(string activity, int number, DataGridView table)
        {
            OleDbDataReader sqlReader = null;
            OleDbCommand command = new OleDbCommand("SELECT Human.Id as Id, Human.Фамилия as Фамилия, Human.Имя as Имя, Human.Отчество as Отчество, Human.Специальность as Специальность, " + activity + ".Результат as Результат FROM " + activity + ", Human WHERE Human.Деятельность="+number +" AND Human.IdActivity=" + activity + ".Id ORDER BY "+activity+".Результат DESC", Data.OleDbConnection); // * считываем все колонки
            try
            {
                sqlReader = command.ExecuteReader();
                while (sqlReader.Read())
                {
                    table.Rows.Add(new string[] {sqlReader["Id"].ToString(),
                                                sqlReader["Фамилия"].ToString(),
                                                sqlReader["Имя"].ToString(),
                                                sqlReader["Отчество"].ToString(),
                                                sqlReader["Специальность"].ToString(),
                                                sqlReader["Результат"].ToString()});
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader != null)
                    sqlReader.Close();
            }
        }

        public static bool IsFilled(List<Control> fields)
        {
            bool isFilled = true;
            foreach (Control field in fields)
            {
                if (field is TextBox)
                {
                    TextBox textBox = field as TextBox;
                    if (String.IsNullOrEmpty(textBox.Text) || IsProbel(textBox.Text))
                    {
                        textBox.BackColor = Color.Pink;
                        isFilled = false;
                    }
                }
                else if (field is MaskedTextBox)
                {
                    MaskedTextBox maskedTextBox = field as MaskedTextBox;
                    if (maskedTextBox.MaskCompleted == false)
                    {
                        maskedTextBox.BackColor = Color.Pink;
                        isFilled = false;
                    }
                }
                else if (field is System.Windows.Forms.CheckBox)
                {
                    System.Windows.Forms.CheckBox checkBox = field as System.Windows.Forms.CheckBox;
                    if (checkBox.Checked == false)
                    {
                        checkBox.BackColor = Color.Pink;
                        isFilled = false;
                    }
                }
                else if (field is DataGridView)
                {
                    DataGridView table = field as DataGridView;
                    for (int i = 0; i < table.RowCount; i++)
                    {
                        if (table.ColumnCount > 2)
                        {
                            int dontKnow;
                            if ((table[1, i].Value == null || IsProbel(table[1, i].Value.ToString())) ||
                                (table[1, i].Value != null && !Int32.TryParse((table[1, i]).Value.ToString(), out dontKnow)))
                            {
                                table[1, i].Style.BackColor = Color.Pink;
                                isFilled = false;
                            }
                            if (table[2, i].Value == null || IsProbel(table[2, i].Value.ToString()))
                            {
                                table[2, i].Style.BackColor = Color.Pink;
                                isFilled = false;
                            }
                        }
                        else
                        {
                            if (table[1, i].Value == null || IsProbel(table[1, i].Value.ToString()))
                            {
                                table[1, i].Style.BackColor = Color.Pink;
                                isFilled = false;
                            }
                        }
                    }
                }
            }
            return isFilled;
        }

        public static void Enter(Control field)
        {
            if (field is DataGridView)
            {
                DataGridView table = field as DataGridView;
                for (int i = 0; i < table.RowCount; i++)
                {
                    if (table.ColumnCount > 2)
                        table[2, i].Style.BackColor = Color.Empty;
                    table[1, i].Style.BackColor = Color.Empty;
                }
            }
            else
                field.BackColor = Color.Empty;
        }
    }

    
}

