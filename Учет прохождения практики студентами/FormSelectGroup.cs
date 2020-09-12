using System;
using System.Windows.Forms;

namespace Учет_прохождения_практики_студентами
{
    public partial class FormSelectGroup : Form
    {
        #region Переменные
        Administration administration = new Administration();
        #endregion

        #region Конструктор
        public FormSelectGroup()
        {
            InitializeComponent();
        }
        #endregion

        #region Загрузка формы
        private void FormSelectGroup_Load(object sender, EventArgs e)
        {
            foreach (string item in administration.RequestDepartmentList())
                comboBox1.Items.Add(item);
            if (comboBox1.Items.Count > 0)
                comboBox1.SelectedIndex = 0;
            comboBox3.SelectedIndex = 0;
        }
        #endregion

        #region Выбор критерия
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cmb = new ComboBox();
            foreach (string item in administration.SearchTeacher(comboBox1.SelectedItem.ToString()))
                cmb.Items.Add(item);
            ((DataGridViewComboBoxColumn)dataGridView1.Columns["Column7"]).DataSource = cmb.Items;
            dataGridView1.Rows.Clear();
            comboBox2.Items.Clear();
            foreach (string item in administration.RequestGroupsList(comboBox1.SelectedItem.ToString()))
                comboBox2.Items.Add(item);
            if (comboBox2.Items.Count > 0)
                comboBox2.SelectedIndex = 0;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            int i = administration.MaxID();
            if (comboBox3.SelectedItem != null && comboBox2.SelectedItem != null)
                if (administration.IsAdd(comboBox2.SelectedItem.ToString(), comboBox3.SelectedItem.ToString()))
                {
                    button1.Text = "Изменить";
                    foreach (Administration.infoStd item in administration.SearchFullWork(comboBox2.SelectedItem.ToString(), comboBox3.SelectedItem.ToString()))
                    {
                        dataGridView1.Rows.Add(item.Value[0]?.ToString() ?? (++i).ToString(), 
                            item.Value[1]?.ToString() ?? "", item.Value[2]?.ToString() ?? "",
                            item.Value[3]?.ToString() ?? "", item.Value[4]?.ToString() ?? "",
                            item.Value[5]?.ToString() ?? "", item.Value[6]?.ToString() ?? "",
                            item.Value[7]?.ToString() ?? "", item.Value[8]?.ToString() ?? "",
                            item.Value[9]?.ToString() ?? "", item.Value[10]?.ToString() ?? "");
                    }
                }
                else
                {
                    if (i != -1)
                    {
                        button1.Text = "Добавить";
                        foreach (string item in administration.SearchNumberStudents(comboBox2.SelectedItem.ToString()))
                        {
                            dataGridView1.Rows.Add(++i, item);
                        }
                    }
                }
        }
        #endregion 

        #region Процесс редактирования
        private void button1_Click(object sender, EventArgs e)
        {
            int request = 0;
            Administration.infoStdString info = new Administration.infoStdString(12);
            if (button1.Text[0] == 'Д')
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    info = new Administration.infoStdString(12);
                    if (IsNull(i)) continue;
                    info.Value[0] = dataGridView1.Rows[i].Cells[0].Value.ToString();
                    info.Value[1] = dataGridView1.Rows[i].Cells[1].Value.ToString();
                    info.Value[2] = dataGridView1.Rows[i].Cells[2].Value.ToString();
                    info.Value[3] = dataGridView1.Rows[i].Cells[3].Value.ToString();
                    info.Value[4] = dataGridView1.Rows[i].Cells[4].Value.ToString();
                    info.Value[5] = dataGridView1.Rows[i].Cells[5].Value.ToString();
                    info.Value[6] = dataGridView1.Rows[i].Cells[6].Value.ToString();
                    info.Value[7] = dataGridView1.Rows[i].Cells[7].Value.ToString();
                    info.Value[8] = dataGridView1.Rows[i].Cells[8].Value.ToString();
                    info.Value[9] = dataGridView1.Rows[i].Cells[9].Value.ToString();
                    info.Value[10] = dataGridView1.Rows[i].Cells[10].Value.ToString();
                    request += administration.AddRequest(info);
                }
            }
            else
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if (IsNull(i)) continue;
                    info.Value[0] = dataGridView1.Rows[i].Cells[0].Value.ToString();
                    info.Value[1] = dataGridView1.Rows[i].Cells[1].Value.ToString();
                    info.Value[2] = dataGridView1.Rows[i].Cells[2].Value.ToString();
                    info.Value[3] = dataGridView1.Rows[i].Cells[3].Value.ToString();
                    info.Value[4] = dataGridView1.Rows[i].Cells[4].Value.ToString();
                    info.Value[5] = dataGridView1.Rows[i].Cells[5].Value.ToString();
                    info.Value[6] = dataGridView1.Rows[i].Cells[6].Value.ToString();
                    info.Value[7] = dataGridView1.Rows[i].Cells[7].Value.ToString();
                    info.Value[8] = dataGridView1.Rows[i].Cells[8].Value.ToString();
                    info.Value[9] = dataGridView1.Rows[i].Cells[9].Value.ToString();
                    info.Value[10] = dataGridView1.Rows[i].Cells[10].Value.ToString();
                    request += administration.ChengeRequest(info);
                }
            }
            string txt = button1.Text[0] == 'Д' ? "добавлено" : "изменено";
            MessageBox.Show($"Было {txt} {request} записей", button1.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            button1.Text = "Изменить";
        }
        #endregion

        #region Проверка заполнения
        private bool IsNull(int rows)
        {
            bool answer = false;
            for (int i = 0; i < 11; i++)
            {
                if (dataGridView1.Rows[rows].Cells[i].Value == null || dataGridView1.Rows[rows].Cells[i].Value.ToString() == String.Empty)
                {
                    answer = true;
                    break;
                }
            }
            return answer;
        }
        #endregion 
    }
}
