using System;
using System.Threading;
using System.Windows.Forms;

namespace Учет_прохождения_практики_студентами
{
    public partial class FormMain : Form
    {
        #region Переменные
        Administration administration = new Administration();
        #endregion

        #region Конструктор
        public FormMain()
        {
            InitializeComponent();
        }
        #endregion 

        #region Загрузка формы
        private void FormMain_Load(object sender, EventArgs e)
        {
            #region Заполнение списков
            foreach (string item in administration.RequestDepartmentList())
            {
                comboBoxDep1.Items.Add(item);
                comboBoxDep2.Items.Add(item);
                comboBoxDep4.Items.Add(item);
                comboBoxDep5.Items.Add(item);
            }
            foreach (string item in administration.RequestManagerList())
                comboBoxOrg3.Items.Add(item);
            #endregion

            #region Назначение индексов
            ReItems(comboBoxDep1);
            ReItems(comboBoxDep2);
            ReItems(comboBoxType2);
            ReItems(comboBoxOrg3);
            ReItems(comboBoxYear3);
            ReItems(comboBoxDep4);
            ReItems(comboBoxYear4);
            ReItems(comboBoxDep5);
            ReItems(comboBoxYear5);            
            #endregion
        }

        private void ReItems(ComboBox comboBox)
        {
            if (comboBox.Items.Count > 0)
            {
                comboBox.Enabled = true;
                comboBox.SelectedIndex = 0;
                comboBoxStd1.Enabled = true;
                button1.Enabled = true;
            }
            else
            {
                comboBox.Items.Clear();
                comboBox.Text = String.Empty;
                comboBox.Enabled = false;
                if (comboBox == comboBoxGr1)
                {
                    comboBoxStd1.Items.Clear();
                    comboBoxStd1.Text = String.Empty;
                    comboBoxStd1.Enabled = false;
                    button1.Enabled = false;
                }
            }
        }

        #endregion

        #region Выбор из списка
        private void ReDepartment(ComboBox comboBoxDep, ComboBox comboBoxGr)
        {
            comboBoxGr.Items.Clear();
            foreach (string item in administration.RequestGroupsList(comboBoxDep.SelectedItem.ToString()))
                comboBoxGr.Items.Add(item);
            ReItems(comboBoxGr);
        }

        private void comboBoxDep1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ReDepartment(comboBoxDep1, comboBoxGr1);
        }

        private void comboBoxGr1_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBoxStd1.Items.Clear();
            foreach (string item in administration.RequestStudentsList(comboBoxGr1.SelectedItem.ToString()))
                comboBoxStd1.Items.Add(item);
            ReItems(comboBoxStd1);
        }

        private void comboBoxStd1_SelectedIndexChanged(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            foreach (string item in administration.SearchByStudent(comboBoxStd1.SelectedItem.ToString()))
                listBox1.Items.Add(item);
        }

        private void comboBoxDep2_SelectedIndexChanged(object sender, EventArgs e)
        {
            ReDepartment(comboBoxDep2, comboBoxGr2);
        }

        private void comboBoxType2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxType2.SelectedItem != null)
            {
                listBox2.Items.Clear();
                foreach (string item in administration.SearchByType(comboBoxType2.SelectedItem.ToString(), comboBoxGr2.SelectedItem.ToString()))
                    listBox2.Items.Add(item);
            }
        }

        private void comboBoxOrg3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxYear3.SelectedItem != null)
            {
                listBox3.Items.Clear();
                foreach (string item in administration.SearchByOrganithetion(comboBoxOrg3.SelectedItem.ToString(), comboBoxYear3.SelectedItem.ToString()))
                    listBox3.Items.Add(item);
            }
        }

        private void comboBoxDep4_SelectedIndexChanged(object sender, EventArgs e)
        {
            ReDepartment(comboBoxDep4, comboBoxGr4);
        }

        private void comboBoxGr4_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxYear4.SelectedItem != null)
            {
                listBox4.Items.Clear();
                foreach (string item in administration.SearchByYear(comboBoxGr4.SelectedItem.ToString(), comboBoxYear4.SelectedItem.ToString()))
                    listBox4.Items.Add(item);
            }
        }

        private void comboBoxDep5_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBoxPr5.Items.Clear();
            foreach (string item in administration.RequestTeachersList(comboBoxDep5.SelectedItem.ToString()))
                comboBoxPr5.Items.Add(item);
            ReItems(comboBoxPr5);
        }

        private void comboBoxPr5_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxYear5.SelectedItem != null)
            {
                listBox5.Items.Clear();
                foreach (string item in administration.SearchByTeacher(comboBoxPr5.SelectedItem.ToString(), comboBoxYear5.SelectedItem.ToString()))
                    listBox5.Items.Add(item);
            }
        }
        #endregion

        #region Формирование отчетов
        private void button1_Click(object sender, EventArgs e)
        {
            Thread thread = new Thread(new ThreadStart(Thread1));
            thread.Start();
        }

        private void Thread1()
        {
            administration.ReportOnStudent(comboBoxStd1.SelectedItem.ToString());
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Thread thread = new Thread(new ThreadStart(Thread2));
            thread.Start();            
        }

        private void Thread2()
        {
            administration.ReportOnType(comboBoxType2.SelectedItem.ToString(), comboBoxGr2.SelectedItem.ToString());
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Thread thread = new Thread(new ThreadStart(Thread3));
            thread.Start();            
        }

        private void Thread3()
        {
            administration.ReportOnOrganithetion(comboBoxOrg3.SelectedItem.ToString(), comboBoxYear3.SelectedItem.ToString());
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Thread thread = new Thread(new ThreadStart(Thread4));
            thread.Start();            
        }

        private void Thread4()
        {
            administration.ReportOnYear(comboBoxGr4.SelectedItem.ToString(), comboBoxYear4.SelectedItem.ToString());
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Thread thread = new Thread(new ThreadStart(Thread5));
            thread.Start();            
        }

        private void Thread5()
        {
            administration.ReportOnTeacher(comboBoxPr5.SelectedItem.ToString(), comboBoxYear5.SelectedItem.ToString());
        }
        #endregion

        #region Добавление или изменение
        private void buttonAdd_Click(object sender, EventArgs e)
        {
            Hide();
            FormSelectGroup form = new FormSelectGroup();
            form.ShowDialog();
            Show();
            comboBoxDep1.SelectedIndex = 0;
            comboBoxDep2.SelectedIndex = 0;
            comboBoxOrg3.SelectedIndex = 0;
            comboBoxDep4.SelectedIndex = 0;
            comboBoxDep5.SelectedIndex = 0;
        }
        #endregion
    }
}
