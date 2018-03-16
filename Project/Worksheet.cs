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


namespace Project
{
    public partial class Worksheet : UserControl
    {
        
        List<Control> fields;

        private static Worksheet instance;
        public static Worksheet Instance
        {
            get
            {
                if (instance == null)
                    instance = new Worksheet();
                return instance;
            }
        }

        public Worksheet()
        {
            InitializeComponent();
            fields = new List<Control>() {surname, name, fatherName, university, birthDay, specialty,
                                          preparation, programm, cours, phone,
                                          homePhone, email, checkBox1};
        }

        private void next_Click(object sender, EventArgs e)
        {
            if (!Data.IsFilled(fields))
            {
                MessageBox.Show("Заполните пустые поля!", "Внимание");
                return;
            }

            Data.human.surname = surname.Text;
            Data.human.name = name.Text;
            Data.human.fatherName = fatherName.Text;
            Data.human.birthDay = birthDay.Text;
            Data.human.university = university.Text;
            Data.human.specialty= specialty.Text;
            Data.human.preparation= preparation.Text;
            Data.human.programm = programm.Text;
            Data.human.cours = Convert.ToInt32(cours.Value);
            Data.human.phone = phone.Text;
            Data.human.homePhone = homePhone.Text;
            Data.human.email = email.Text;
            if (radioButton1.Checked)
                Data.human.activity = 1;
            else if (radioButton2.Checked)
                Data.human.activity = 2;
            else if (radioButton3.Checked)
                Data.human.activity = 3;
            else if (radioButton4.Checked)
                Data.human.activity = 4;
            else
                Data.human.activity = 5;
            Data.human.SetInWord();

            if (radioButton1.Checked)
            {
                if (!Data.panel.Controls.Contains(EducationActivity.Instance)) //если нет в массиве то добавляем
                {
                    Data.panel.Controls.Add(EducationActivity.Instance);
                    EducationActivity.Instance.Dock = DockStyle.Fill;
                    EducationActivity.Instance.BringToFront();
                }
                else
                    EducationActivity.Instance.BringToFront();
            }
            else if (radioButton2.Checked)
            {
                if (!Data.panel.Controls.Contains(ScientificResearchAcivity.Instance)) //если нет в массиве то добавляем
                {
                    Data.panel.Controls.Add(ScientificResearchAcivity.Instance);
                    ScientificResearchAcivity.Instance.Dock = DockStyle.Fill;
                    ScientificResearchAcivity.Instance.BringToFront();
                }
                else
                    ScientificResearchAcivity.Instance.BringToFront();
            }
            else if (radioButton3.Checked)
            {
                if (!Data.panel.Controls.Contains(СulturalСreativeActivity.Instance)) //если нет в массиве то добавляем
                {
                    Data.panel.Controls.Add(СulturalСreativeActivity.Instance);
                    СulturalСreativeActivity.Instance.Dock = DockStyle.Fill;
                    СulturalСreativeActivity.Instance.BringToFront();
                }
                else
                    СulturalСreativeActivity.Instance.BringToFront();
            }
            else if (radioButton4.Checked)
            {
                if (!Data.panel.Controls.Contains(PublicActivity.Instance)) //если нет в массиве то добавляем
                {
                    Data.panel.Controls.Add(PublicActivity.Instance);
                    PublicActivity.Instance.Dock = DockStyle.Fill;
                    PublicActivity.Instance.BringToFront();
                }
                else
                    PublicActivity.Instance.BringToFront();
            }
            else
            {
                if (!Data.panel.Controls.Contains(SportActivity.Instance)) //если нет в массиве то добавляем
                {
                    Data.panel.Controls.Add(SportActivity.Instance);
                    SportActivity.Instance.Dock = DockStyle.Fill;
                    SportActivity.Instance.BringToFront();
                }
                else
                    SportActivity.Instance.BringToFront();
            }
        }

        private void field_Enter(object sender, EventArgs e)
        {
            Control control = sender as Control;
            control.BackColor = Color.Empty;
        }
    }
}
