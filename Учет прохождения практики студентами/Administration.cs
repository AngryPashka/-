using System;
using System.Collections.Generic;
using System.Data.OleDb;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Windows.Forms;

namespace Учет_прохождения_практики_студентами
{
    class Administration
    {
        #region Переменные
        OleDbCommand command;
        OleDbDataReader reader;
        #endregion

        #region Конструктор
        public Administration() { }
        #endregion

        #region Аутентификация
        public bool Authentication(string name, string password)
        {
            OleDbConnection connection = new OleDbConnection(@"Provider = Microsoft.Jet.OLEDB.4.0; Data Source = " + Program.UsersBD);
            connection.Open();
            OleDbCommand command = new OleDbCommand($"SELECT Users.ID FROM Users WHERE Users.Login = '{name}' AND Users.Password = {password}", connection);
            OleDbDataReader reader = command.ExecuteReader();
            bool answer = reader.HasRows;
            reader.Dispose();
            reader.Close();
            command.Dispose();
            connection.Close();
            connection.Dispose();
            return answer;
        }
        #endregion

        #region Запросы на заполнение
        public List<string> RequestDepartmentList()
        {
            List<string> answer = new List<string>();
            OleDbConnection connection = new OleDbConnection(@"Provider =  Microsoft.Jet.OLEDB.4.0; Data Source = " + Program.PracticeBD);
            connection.Open();
            command = new OleDbCommand("SELECT Кафедра.Код FROM Кафедра", connection);
            reader = command.ExecuteReader();
            if (reader.HasRows)
                while (reader.Read())
                    answer.Add(reader.GetString(0));
            reader.Dispose();
            reader.Close();
            command.Dispose();
            connection.Close();
            connection.Dispose();
            return answer;
        }

        public List<string> RequestManagerList()
        {
            OleDbConnection connection = new OleDbConnection(@"Provider =  Microsoft.Jet.OLEDB.4.0; Data Source = " + Program.PracticeBD);
            List<string> answer = new List<string>();
            connection.Open();
            command = new OleDbCommand("SELECT Практика.[Место прохождения] FROM Практика GROUP BY Практика.[Место прохождения] ORDER BY Практика.[Место прохождения]", connection);
            reader = command.ExecuteReader();
            if (reader.HasRows)
                while (reader.Read())
                    answer.Add(reader.GetString(0));
            reader.Dispose();
            reader.Close();
            connection.Close();
            connection.Dispose();
            return answer;
        }

        public List<string> RequestGroupsList(string name)
        {
            OleDbConnection connection = new OleDbConnection(@"Provider =  Microsoft.Jet.OLEDB.4.0; Data Source = " + Program.PracticeBD);
            List<string> answer = new List<string>();
            connection.Open();
            command = new OleDbCommand($"SELECT [Учебная группа].Код FROM Кафедра INNER JOIN [Учебная группа] ON Кафедра.Код = [Учебная группа].[Код кафедры] WHERE Кафедра.Код = '{name}'", connection);
            reader = command.ExecuteReader();
            if (reader.HasRows)
                while (reader.Read())
                    answer.Add(reader.GetString(0));
            reader.Dispose();
            reader.Close();
            command.Dispose();
            connection.Close();
            connection.Dispose();
            return answer;
        }

        public List<string> RequestStudentsList(string name)
        {
            OleDbConnection connection = new OleDbConnection(@"Provider =  Microsoft.Jet.OLEDB.4.0; Data Source = " + Program.PracticeBD);
            List<string> answer = new List<string>();
            connection.Open();
            command = new OleDbCommand($"SELECT Студент.[Номер зачетки] FROM [Учебная группа] INNER JOIN Студент ON [Учебная группа].Код = Студент.[Код учебной группы] WHERE [Учебная группа].Код = '{name}'", connection);
            reader = command.ExecuteReader();
            if (reader.HasRows)
                while (reader.Read())
                    answer.Add(reader.GetString(0));
            reader.Dispose();
            reader.Close();
            command.Dispose();
            connection.Close();
            connection.Dispose();
            return answer;
        }

        public List<string> RequestTeachersList(string name)
        {
            OleDbConnection connection = new OleDbConnection(@"Provider =  Microsoft.Jet.OLEDB.4.0; Data Source = " + Program.PracticeBD);
            List<string> answer = new List<string>();
            connection.Open();
            command = new OleDbCommand($"SELECT Преподаватель.ФИО FROM Кафедра INNER JOIN Преподаватель ON Кафедра.Код = Преподаватель.[Код кафедры] WHERE Кафедра.Код = '{name}'", connection);
            reader = command.ExecuteReader();
            if (reader.HasRows)
                while (reader.Read())
                    answer.Add(reader.GetString(0));
            reader.Dispose();
            reader.Close();
            command.Dispose();
            connection.Close();
            connection.Dispose();
            return answer;
        }

        public List<string> RequestGroups(string name)
        {
            OleDbConnection connection = new OleDbConnection(@"Provider =  Microsoft.Jet.OLEDB.4.0; Data Source = " + Program.PracticeBD);
            List<string> answer = new List<string>();
            connection.Open();
            command = new OleDbCommand($"SELECT [Учебная группа].Код FROM Кафедра WHERE Кафедра.Код = '{name}'", connection);
            reader = command.ExecuteReader();
            if (reader.HasRows)
                while (reader.Read())
                    answer.Add(reader.GetString(0));
            reader.Dispose();
            reader.Close();
            command.Dispose();
            connection.Close();
            connection.Dispose();
            return answer;
        }

        #endregion

        #region Запросы на наполнение
        public List<string> SearchByStudent(string Number)
        {
            OleDbConnection connection = new OleDbConnection(@"Provider =  Microsoft.Jet.OLEDB.4.0; Data Source = " + Program.PracticeBD);
            List<string> answer = new List<string>();
            connection.Open();
            command = new OleDbCommand($"SELECT Практика.[Название практики], Практика.[Место прохождения], Практика.[Дата начала], Практика.[Дата окончания], Преподаватель.ФИО, Практика.[Руководитель от предприятия], Практика.Оценка, Практика.[Дата выставления] FROM Преподаватель INNER JOIN Практика ON Преподаватель.ID = Практика.[Руководитель от вуза] WHERE Практика.[Зачетная книжка]= '{Number}'", connection);
            reader = command.ExecuteReader();
            if (reader.HasRows)
                while (reader.Read())
                    answer.Add(String.Format("{0} в {1} с {2} по {3}, руководитель: от вуза - {4}, от пред. - {5}, оценка: {6} ({7})",
                        reader.GetString(0), reader.GetString(1),
                        Convert.ToDateTime(reader.GetValue(2)).ToShortDateString(),
                        Convert.ToDateTime(reader.GetValue(3)).ToShortDateString(),
                        reader.GetValue(4), reader.GetString(5), reader.GetValue(6),
                        Convert.ToDateTime(reader.GetValue(7)).ToShortDateString()));
            reader.Dispose();
            reader.Close();
            command.Dispose();
            connection.Close();
            connection.Dispose();
            return answer;
        }

        public List<string> SearchByType(string Type, string Group)
        {
            OleDbConnection connection = new OleDbConnection(@"Provider =  Microsoft.Jet.OLEDB.4.0; Data Source = " + Program.PracticeBD);
            List<string> answer = new List<string>();
            connection.Open();
            command = new OleDbCommand($"SELECT Практика.[Зачетная книжка], Практика.[Место прохождения], Практика.[Дата начала], Практика.[Дата окончания], Преподаватель.ФИО, Практика.[Руководитель от предприятия], Практика.Оценка, Практика.[Дата выставления] FROM Студент INNER JOIN (Преподаватель INNER JOIN Практика ON Преподаватель.ID = Практика.[Руководитель от вуза]) ON Студент.[Номер зачетки] = Практика.[Зачетная книжка] WHERE Практика.[Название практики]= '{Type}' AND Студент.[Код учебной группы]= '{Group}'", connection);
            reader = command.ExecuteReader();
            if (reader.HasRows)
                while (reader.Read())
                    answer.Add(String.Format("Студент {0} прохоил практику в {1} с {2} по {3}, руководитель: от вуза - {4}, от пред. - {5}, оценка: {6} ({7})",
                        reader.GetString(0), reader.GetString(1),
                        Convert.ToDateTime(reader.GetValue(2)).ToShortDateString(),
                        Convert.ToDateTime(reader.GetValue(3)).ToShortDateString(),
                        reader.GetValue(4), reader.GetString(5), reader.GetValue(6),
                        Convert.ToDateTime(reader.GetValue(7)).ToShortDateString()));
            reader.Dispose();
            reader.Close();
            command.Dispose();
            connection.Close();
            connection.Dispose();
            return answer;
        }

        public List<string> SearchByOrganithetion(string Org, string Year)
        {
            OleDbConnection connection = new OleDbConnection(@"Provider =  Microsoft.Jet.OLEDB.4.0; Data Source = " + Program.PracticeBD);
            List<string> answer = new List<string>();
            connection.Open();
            if (Year == "Все")
            {
                command = new OleDbCommand($"SELECT Практика.[Зачетная книжка], Практика.[Название практики], Практика.[Дата начала], Практика.[Дата окончания], Преподаватель.ФИО, Практика.[Руководитель от предприятия], Практика.Оценка, Практика.[Дата выставления] FROM Преподаватель INNER JOIN Практика ON Преподаватель.ID = Практика.[Руководитель от вуза] WHERE Практика.[Место прохождения]= '{Org}'", connection);
            }
            else
            {
                command = new OleDbCommand($"SELECT Практика.[Зачетная книжка], Практика.[Название практики], Практика.[Дата начала], Практика.[Дата окончания], Преподаватель.ФИО, Практика.[Руководитель от предприятия], Практика.Оценка, Практика.[Дата выставления] FROM Преподаватель INNER JOIN Практика ON Преподаватель.ID = Практика.[Руководитель от вуза] WHERE Практика.[Место прохождения] = '{Org}' AND Year([Практика].[Дата начала])= '{Year}'", connection);
            }
            reader = command.ExecuteReader();
            if (reader.HasRows)
                while (reader.Read())
                    answer.Add(String.Format("Студент {0} прохоил {1} с {2} по {3}, руководитель: от вуза - {4}, от пред. - {5}, оценка: {6} ({7})",
                        reader.GetString(0), reader.GetString(1),
                        Convert.ToDateTime(reader.GetValue(2)).ToShortDateString(),
                        Convert.ToDateTime(reader.GetValue(3)).ToShortDateString(),
                        reader.GetValue(4), reader.GetString(5), reader.GetValue(6),
                        Convert.ToDateTime(reader.GetValue(7)).ToShortDateString()));
            reader.Dispose();
            reader.Close();
            command.Dispose();
            connection.Close();
            connection.Dispose();
            return answer;
        }

        public List<string> SearchByYear(string Group, string Year)
        {
            OleDbConnection connection = new OleDbConnection(@"Provider =  Microsoft.Jet.OLEDB.4.0; Data Source = " + Program.PracticeBD);
            List<string> answer = new List<string>();
            connection.Open();
            if (Year == "Все")
            {
                command = new OleDbCommand($"SELECT Практика.[Зачетная книжка], Практика.[Название практики], Практика.[Место прохождения], Практика.[Дата начала], Практика.[Дата окончания], Преподаватель.ФИО, Практика.[Руководитель от предприятия], Практика.Оценка, Практика.[Дата выставления] FROM Студент INNER JOIN (Преподаватель INNER JOIN Практика ON Преподаватель.ID = Практика.[Руководитель от вуза]) ON Студент.[Номер зачетки] = Практика.[Зачетная книжка] WHERE Студент.[Код учебной группы]= '{Group}'", connection);
            }
            else
            {
                command = new OleDbCommand($"SELECT Практика.[Зачетная книжка], Практика.[Название практики], Практика.[Место прохождения], Практика.[Дата начала], Практика.[Дата окончания], Преподаватель.ФИО, Практика.[Руководитель от предприятия], Практика.Оценка, Практика.[Дата выставления] FROM Студент INNER JOIN (Преподаватель INNER JOIN Практика ON Преподаватель.ID = Практика.[Руководитель от вуза]) ON Студент.[Номер зачетки] = Практика.[Зачетная книжка] WHERE YEAR(Практика.[Дата начала])= '{Year}' AND Студент.[Код учебной группы]= '{Group}'", connection);
            }
            reader = command.ExecuteReader();
            if (reader.HasRows)
                while (reader.Read())
                    answer.Add(String.Format("Студент {0} прохоил {1} в {2} с {3} по {4}, руководитель: от вуза - {5}, от пред. - {6}, оценка: {7} ({8})",
                        reader.GetString(0), reader.GetString(1), reader.GetString(2),
                        Convert.ToDateTime(reader.GetValue(3)).ToShortDateString(),
                        Convert.ToDateTime(reader.GetValue(4)).ToShortDateString(),
                        reader.GetString(5), reader.GetString(6), reader.GetValue(7),
                        Convert.ToDateTime(reader.GetValue(8)).ToShortDateString()));
            reader.Dispose();
            reader.Close();
            command.Dispose();
            connection.Close();
            connection.Dispose();
            return answer;
        }

        public List<string> SearchByTeacher(string Name, string Year)
        {
            OleDbConnection connection = new OleDbConnection(@"Provider =  Microsoft.Jet.OLEDB.4.0; Data Source = " + Program.PracticeBD);
            List<string> answer = new List<string>();
            connection.Open();
            if (Year == "Все")
            {
                command = new OleDbCommand($"SELECT Практика.[Зачетная книжка], Практика.[Название практики], Практика.[Место прохождения], Практика.[Дата начала], Практика.[Дата окончания], Практика.[Руководитель от предприятия], Практика.Оценка, Практика.[Дата выставления], Практика.[Дата начала] FROM Студент INNER JOIN (Преподаватель INNER JOIN Практика ON Преподаватель.ID = Практика.[Руководитель от вуза]) ON Студент.[Номер зачетки] = Практика.[Зачетная книжка] WHERE Преподаватель.ФИО= '{Name}'", connection);
            }
            else
            {
                command = new OleDbCommand($"SELECT Практика.[Зачетная книжка], Практика.[Название практики], Практика.[Место прохождения], Практика.[Дата начала], Практика.[Дата окончания], Практика.[Руководитель от предприятия], Практика.Оценка, Практика.[Дата выставления], Практика.[Дата начала] FROM Студент INNER JOIN (Преподаватель INNER JOIN Практика ON Преподаватель.ID = Практика.[Руководитель от вуза]) ON Студент.[Номер зачетки] = Практика.[Зачетная книжка] WHERE Преподаватель.ФИО= '{Name}' AND YEAR(Практика.[Дата начала])= '{Year}'", connection);
            }
            reader = command.ExecuteReader();
            if (reader.HasRows)
                while (reader.Read())
                    answer.Add(String.Format("Студент {0} прохоил {1} в {2} с {3} по {4}, руководитель от пред. - {5}, оценка: {6} ({7})",
                        reader.GetString(0), reader.GetString(1), reader.GetString(2),
                        Convert.ToDateTime(reader.GetValue(3)).ToShortDateString(),
                        Convert.ToDateTime(reader.GetValue(4)).ToShortDateString(),
                        reader.GetString(5), reader.GetValue(6),
                        Convert.ToDateTime(reader.GetValue(7)).ToShortDateString()));
            reader.Dispose();
            reader.Close();
            command.Dispose();
            connection.Close();
            connection.Dispose();
            return answer;
        }

        #endregion

        #region Редактирование БД
        public List<string> SearchNumberStudents(string Group)
        {
            OleDbConnection connection = new OleDbConnection(@"Provider =  Microsoft.Jet.OLEDB.4.0; Data Source = " + Program.PracticeBD);
            List<string> answer = new List<string>();
            connection.Open();
            command = new OleDbCommand($"SELECT Студент.[Номер зачетки] FROM Студент WHERE Студент.[Код учебной группы] = '{Group}'", connection);
            reader = command.ExecuteReader();
            if (reader.HasRows)
                while (reader.Read()) answer.Add(reader.GetString(0));
            reader.Dispose();
            reader.Close();
            command.Dispose();
            connection.Close();
            connection.Dispose();
            return answer;
        }

        public List<string> SearchTeacher(string Dep)
        {
            OleDbConnection connection = new OleDbConnection(@"Provider =  Microsoft.Jet.OLEDB.4.0; Data Source = " + Program.PracticeBD);
            List<string> answer = new List<string>();
            connection.Open();
            command = new OleDbCommand($"SELECT Преподаватель.ФИО FROM Преподаватель WHERE Преподаватель.[Код кафедры]= '{Dep}'", connection);
            reader = command.ExecuteReader();
            if (reader.HasRows)
                while (reader.Read())
                    answer.Add(reader.GetString(0));
            reader.Dispose();
            reader.Close();
            command.Dispose();
            connection.Close();
            connection.Dispose();
            return answer;
        }

        public bool IsAdd(string Group, string Year)
        {
            OleDbConnection connection = new OleDbConnection(@"Provider =  Microsoft.Jet.OLEDB.4.0; Data Source = " + Program.PracticeBD);         
            connection.Open();
            command = new OleDbCommand($"SELECT TOP 1 Студент.[Код учебной группы] FROM Студент INNER JOIN (Преподаватель INNER JOIN Практика ON Преподаватель.ID = Практика.[Руководитель от вуза]) ON Студент.[Номер зачетки] = Практика.[Зачетная книжка] WHERE Студент.[Код учебной группы]= '{Group}' AND YEAR(Практика.[Дата начала])= {Year}", connection);
            reader = command.ExecuteReader();
            bool answer = reader.HasRows;
            reader.Dispose();
            reader.Close();
            command.Dispose();
            connection.Close();
            connection.Dispose();
            return answer;
        }

        public int MaxID()
        {
            OleDbConnection connection = new OleDbConnection(@"Provider =  Microsoft.Jet.OLEDB.4.0; Data Source = " + Program.PracticeBD);
            connection.Open();
            command = new OleDbCommand($"SELECT TOP 1 Практика.ID FROM Практика ORDER BY Практика.ID DESC", connection);
            reader = command.ExecuteReader();
            int answer = -1;
            if (reader.HasRows)
            {
                reader.Read();
                answer = reader.GetInt32(0);
            }
            reader.Dispose();
            reader.Close();
            command.Dispose();
            connection.Close();
            connection.Dispose();
            return answer;
        }

        public List<infoStd> SearchFullWork(string Group, string Year)
        {           
            List<infoStd> list = new List<infoStd>();
            infoStd info;
            info = new infoStd(11);
            foreach (string item in SearchNumberStudents(Group))
            {
                info = new infoStd(11);
                info.Value[1] = item;
                list.Add(info);
            }
            OleDbConnection connection = new OleDbConnection(@"Provider =  Microsoft.Jet.OLEDB.4.0; Data Source = " + Program.PracticeBD);
            connection.Open();
            command = new OleDbCommand($"SELECT Практика.ID, Практика.[Зачетная книжка], Практика.[Название практики], Практика.[Место прохождения], Практика.[Дата начала], Практика.[Дата окончания], Преподаватель.ФИО, Практика.[Руководитель от предприятия], Практика.[Отзыв руководителя], Практика.Оценка, Практика.[Дата выставления] FROM Студент INNER JOIN (Преподаватель INNER JOIN Практика ON Преподаватель.ID = Практика.[Руководитель от вуза]) ON Студент.[Номер зачетки] = Практика.[Зачетная книжка] WHERE Студент.[Код учебной группы]= '{Group}' AND Year([Практика].[Дата начала]) = {Year}", connection);
            try
            {
                reader = command.ExecuteReader();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            int i = 0;
            if (reader.HasRows)
                while (reader.Read())
                {
                    info = new infoStd(11);
                    info.Value[0] = reader.GetValue(0);
                    info.Value[1] = reader.GetValue(1);
                    info.Value[2] = reader.GetValue(2);
                    info.Value[3] = reader.GetValue(3);
                    info.Value[4] = Convert.ToDateTime(reader.GetValue(4)).ToShortDateString();
                    info.Value[5] = Convert.ToDateTime(reader.GetValue(5)).ToShortDateString();
                    info.Value[6] = reader.GetValue(6);
                    info.Value[7] = reader.GetValue(7);
                    info.Value[8] = reader.GetValue(8);
                    info.Value[9] = reader.GetValue(9);
                    info.Value[10] = Convert.ToDateTime(reader.GetValue(10)).ToShortDateString();
                    list[i] = info;
                    i++;
                    if (i >= list.Count) break;
                }            
            reader.Dispose();
            reader.Close();
            command.Dispose();
            connection.Close();
            connection.Dispose();
            return list;
        }

        public int AddRequest(infoStdString info)
        {
            OleDbConnection connection = new OleDbConnection(@"Provider = Microsoft.Jet.OLEDB.4.0; Data Source = " + Program.PracticeBD);
            connection.Open();
            command = new OleDbCommand($"INSERT INTO Практика VALUES " + 
                $"('{info.Value[0]}', '{info.Value[1]}'," + 
                $" '{info.Value[2]}', '{info.Value[3]}'," + 
                $" '{info.Value[4]}', '{info.Value[5]}',"+ 
                $" '{IDTeacher(info.Value[6].ToString())}', '{info.Value[7]}'," + 
                $" '{info.Value[8]}', '{info.Value[9]}'," +
                $" '{info.Value[10]}')", connection);
            int answer = command.ExecuteNonQuery();
            reader.Dispose();
            reader.Close();
            command.Dispose();
            connection.Close();
            connection.Dispose();
            return answer;
        }

        public int ChengeRequest(infoStdString info)
        {
            int answer = 0;
            if (Convert.ToInt32(info.Value[0]) > MaxID())
                answer = AddRequest(info);
            else
            {
                OleDbConnection connection = new OleDbConnection(@"Provider =  Microsoft.Jet.OLEDB.4.0; Data Source = " + Program.PracticeBD);
                connection.Open();
                command = new OleDbCommand($"UPDATE Практика SET Практика.[Зачетная книжка] = '{info.Value[1]}', Практика.[Название практики] = '{info.Value[2]}', Практика.[Место прохождения] = '{info.Value[3]}'," +
               $" Практика.[Дата начала] = '{info.Value[4]}', Практика.[Дата окончания] = '{info.Value[5]}', Практика.[Руководитель от вуза] = {IDTeacher(info.Value[6].ToString())}, Практика.[Руководитель от предприятия] = '{info.Value[7]}'," +
               $" Практика.[Отзыв руководителя] = '{info.Value[8]}', Практика.Оценка = '{info.Value[9]}', Практика.[Дата выставления] = '{info.Value[10]}' WHERE Практика.ID = @ID", connection);
                command.Parameters.Add(new OleDbParameter("@ID", info.Value[0]));
                answer = command.ExecuteNonQuery();
                command.Dispose();
                connection.Close();
                connection.Dispose();
            }            
            return answer;
        }

        public struct infoStdString
        {
            public string[] Value;
            public infoStdString(int i)
            {
                Value = new string[i];
            }
        }

        private int IDTeacher(string Name)
        {
            OleDbConnection connection = new OleDbConnection(@"Provider =  Microsoft.Jet.OLEDB.4.0; Data Source = " + Program.PracticeBD);
            connection.Open();
            command = new OleDbCommand($"SELECT Преподаватель.ID FROM Преподаватель WHERE Преподаватель.ФИО = '{Name}'", connection);
            reader = command.ExecuteReader();
            int answer = -1;
            if (reader.HasRows)
            {
                reader.Read();
                answer = reader.GetInt32(0);
            }
            reader.Dispose();
            reader.Close();
            command.Dispose();
            connection.Close();
            connection.Dispose();
            return answer;
        }
        #endregion

        #region Формаривание отчетов

        public struct infoStd
        {
            public object[] Value;
            public infoStd(int i)
            {
                Value = new object[i];
            }
        }
        public void ReportOnStudent(string Number)
        {
            object oMissing = System.Reflection.Missing.Value;
            object oEndOfDoc = "\\endofdoc";
            Word.Application oWord = new Word.Application();
            Word.Document oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
            ref oMissing, ref oMissing);
            oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
            oWord.Visible = true;            
            object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            Word.Paragraph oPara1;
            oPara1 = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara1.Range.Text = "Отчет по практикам студента №з/кн - " + Number;
            oPara1.Range.Font.Bold = 1;
            oPara1.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            oPara1.Format.SpaceAfter = 24;
            oPara1.Range.InsertParagraphAfter();
            OleDbConnection connection = new OleDbConnection(@"Provider =  Microsoft.Jet.OLEDB.4.0; Data Source = " + Program.PracticeBD);
            List<infoStd> list = new List<infoStd>();
            connection.Open();
            infoStd info;
            command = new OleDbCommand($"SELECT Практика.[Название практики], Практика.[Место прохождения], Практика.[Дата начала], Практика.[Дата окончания], Преподаватель.ФИО, Практика.[Руководитель от предприятия], Практика.Оценка, Практика.[Дата выставления] FROM Преподаватель INNER JOIN Практика ON Преподаватель.ID = Практика.[Руководитель от вуза] WHERE Практика.[Зачетная книжка]= '{Number}'", connection);
            reader = command.ExecuteReader();
            if (reader.HasRows)
                while (reader.Read())
                {
                    info = new infoStd(8);
                    info.Value[0] = reader.GetString(0);
                    info.Value[1] = reader.GetString(1);
                    info.Value[2] = Convert.ToDateTime(reader.GetValue(2)).ToShortDateString();
                    info.Value[3] = Convert.ToDateTime(reader.GetValue(3)).ToShortDateString();
                    info.Value[4] = reader.GetValue(4);
                    info.Value[5] = reader.GetString(5);
                    info.Value[6] = reader.GetValue(6);
                    info.Value[7] = Convert.ToDateTime(reader.GetValue(7)).ToShortDateString();
                    list.Add(info);
                }
            Word.Table oTable;
            Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oTable = oDoc.Tables.Add(wrdRng, list.Count + 2, 8, ref oMissing, ref oMissing);
            oTable.Range.ParagraphFormat.SpaceAfter = 6;
            oTable.Range.Font.Bold = 0;            
            oTable.Range.Font.Size = 10;
            oTable.Cell(1, 1).Range.Text = "Название практики";
            oTable.Cell(1, 2).Range.Text = "Место";
            oTable.Cell(1, 3).Range.Text = "Дата начала";
            oTable.Cell(1, 4).Range.Text = "Дата окончания";
            oTable.Cell(1, 5).Range.Text = "Рук-ль от вуза";
            oTable.Cell(1, 6).Range.Text = "Рук-ть от пред.";
            oTable.Cell(1, 7).Range.Text = "Оценка";
            oTable.Cell(1, 8).Range.Text = "Дата выст-я";
            for (int i = 0; i < 9; i++)
            {
                oTable.Cell(1, i).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleInset;
            }
            double dR = 0;
            for (int i = 2; i <= list.Count + 1; i++)
                for (int j = 1; j <= 8; j++)
                {
                    oTable.Cell(i, j).Range.Text = list[i - 2].Value[j - 1].ToString();
                    oTable.Cell(i, j).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleInset;
                    if (j == 7)
                    {
                        switch (list[i - 2].Value[j - 1].ToString())
                        {
                            case "Отл": dR += 5; break;
                            case "Хор": dR += 4; break;
                            case "Удв": dR += 3; break;
                            case "Неуд": dR += 2; break;
                        }                        
                    }
                }

            oTable.Cell(list.Count + 2, 6).Range.Text = "Средняя оценка:";
            oTable.Cell(list.Count + 2, 6).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleInset;

            oTable.Cell(list.Count + 2, 7).Range.Text = (dR/list.Count).ToString("0.000");
            oTable.Cell(list.Count + 2, 7).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleInset;
            reader.Dispose();
            reader.Close();
            command.Dispose();
            connection.Close();
            connection.Dispose();
        }

        public void ReportOnType(string Type, string Group)
        {
            object oMissing = Missing.Value;
            object oEndOfDoc = "\\endofdoc";
            Word._Application oWord = new Word.Application();
            Word._Document oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
            ref oMissing, ref oMissing);
            oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
            oWord.Visible = true;
            object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            Word.Paragraph oPara1;
            oPara1 = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara1.Range.Text = String.Format("Отчет по {0} группы {1}", Type, Group);
            oPara1.Range.Font.Bold = 1;
            oPara1.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            oPara1.Format.SpaceAfter = 24;
            oPara1.Range.InsertParagraphAfter();
            OleDbConnection connection = new OleDbConnection(@"Provider =  Microsoft.Jet.OLEDB.4.0; Data Source = " + Program.PracticeBD);
            List<infoStd> list = new List<infoStd>();
            connection.Open();
            infoStd info;
            command = new OleDbCommand($"SELECT Практика.[Зачетная книжка], Практика.[Место прохождения], Практика.[Дата начала], Практика.[Дата окончания], Преподаватель.ФИО, Практика.[Руководитель от предприятия], Практика.Оценка, Практика.[Дата выставления] FROM Студент INNER JOIN (Преподаватель INNER JOIN Практика ON Преподаватель.ID = Практика.[Руководитель от вуза]) ON Студент.[Номер зачетки] = Практика.[Зачетная книжка] WHERE Практика.[Название практики]= '{Type}' AND Студент.[Код учебной группы]= '{Group}'", connection);
            reader = command.ExecuteReader();
            if (reader.HasRows)
                while (reader.Read())
                {
                    info = new infoStd(8);
                    info.Value[0] = reader.GetString(0);
                    info.Value[1] = reader.GetString(1);
                    info.Value[2] = Convert.ToDateTime(reader.GetValue(2)).ToShortDateString();
                    info.Value[3] = Convert.ToDateTime(reader.GetValue(3)).ToShortDateString();
                    info.Value[4] = reader.GetValue(4);
                    info.Value[5] = reader.GetString(5);
                    info.Value[6] = reader.GetValue(6);
                    info.Value[7] = Convert.ToDateTime(reader.GetValue(7)).ToShortDateString();
                    list.Add(info);
                }
            Word.Table oTable;
            Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oTable = oDoc.Tables.Add(wrdRng, list.Count + 2, 8, ref oMissing, ref oMissing);
            oTable.Range.ParagraphFormat.SpaceAfter = 6;
            oTable.Range.Font.Bold = 0;
            oTable.Range.Font.Size = 8;
            oTable.Cell(1, 1).Range.Text = "Зачетная книжка";
            oTable.Cell(1, 2).Range.Text = "Место";
            oTable.Cell(1, 3).Range.Text = "Дата начала";
            oTable.Cell(1, 4).Range.Text = "Дата окончания";
            oTable.Cell(1, 5).Range.Text = "Рук-ль от вуза";
            oTable.Cell(1, 6).Range.Text = "Рук-ть от пред.";
            oTable.Cell(1, 7).Range.Text = "Оценка";
            oTable.Cell(1, 8).Range.Text = "Дата выст-я";
            for (int i = 0; i < 9; i++)
            {
                oTable.Cell(1, i).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleInset;
            }
            double dR = 0;
            for (int i = 2; i <= list.Count + 1; i++)
                for (int j = 1; j <= 8; j++)
                {
                    oTable.Cell(i, j).Range.Text = list[i - 2].Value[j - 1].ToString();
                    oTable.Cell(i, j).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleInset;
                    if (j == 7)
                    {
                        switch (list[i - 2].Value[j - 1].ToString())
                        {
                            case "Отл": dR += 5; break;
                            case "Хор": dR += 4; break;
                            case "Удв": dR += 3; break;
                            case "Неуд": dR += 2; break;
                        }
                    }
                }

            oTable.Cell(list.Count + 2, 6).Range.Text = "Средняя оценка:";
            oTable.Cell(list.Count + 2, 6).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleInset;

            oTable.Cell(list.Count + 2, 7).Range.Text = (dR / list.Count).ToString("0.000");
            oTable.Cell(list.Count + 2, 7).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleInset;
            reader.Dispose();
            reader.Close();
            command.Dispose();
            connection.Close();
            connection.Dispose();
        }

        public void ReportOnOrganithetion(string Org, string Year)
        {
            object oMissing = Missing.Value;
            object oEndOfDoc = "\\endofdoc";
            Word.Application oWord = new Word.Application();
            Word.Document oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
            ref oMissing, ref oMissing);
            oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
            oWord.Visible = true;
            try
            {
                object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                OleDbConnection connection = new OleDbConnection(@"Provider =  Microsoft.Jet.OLEDB.4.0; Data Source = " + Program.PracticeBD);
                List<infoStd> list = new List<infoStd>();
                connection.Open();
                infoStd info;
                Word.Paragraph oPara1;
                oPara1 = oDoc.Content.Paragraphs.Add(ref oMissing);
                if (Year == "Все")
                {
                    oPara1.Range.Text = String.Format("Отчет по студентам проходивших практику в {0}", Org);
                    command = new OleDbCommand($"SELECT Практика.[Зачетная книжка], Практика.[Название практики], Практика.[Дата начала], Практика.[Дата окончания], Преподаватель.ФИО, Практика.[Руководитель от предприятия], Практика.Оценка, Практика.[Дата выставления] FROM Преподаватель INNER JOIN Практика ON Преподаватель.ID = Практика.[Руководитель от вуза] WHERE Практика.[Место прохождения]= '{Org}'", connection);
                }
                else
                {
                    oPara1.Range.Text = String.Format("Отчет по студентам проходивших практику в {0} в {1} г.", Org, Year);
                    command = new OleDbCommand($"SELECT Практика.[Зачетная книжка], Практика.[Название практики], Практика.[Дата начала], Практика.[Дата окончания], Преподаватель.ФИО, Практика.[Руководитель от предприятия], Практика.Оценка, Практика.[Дата выставления] FROM Преподаватель INNER JOIN Практика ON Преподаватель.ID = Практика.[Руководитель от вуза] WHERE Практика.[Место прохождения] = '{Org}' AND Year([Практика].[Дата начала])= '{Year}'", connection);
                }
                oPara1.Range.Font.Bold = 1;
                oPara1.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                oPara1.Format.SpaceAfter = 24;
                oPara1.Range.InsertParagraphAfter();
                reader = command.ExecuteReader();
                if (reader.HasRows)
                    while (reader.Read())
                    {
                        info = new infoStd(8);
                        info.Value[0] = reader.GetString(0);
                        info.Value[1] = reader.GetString(1);
                        info.Value[2] = Convert.ToDateTime(reader.GetValue(2)).ToShortDateString();
                        info.Value[3] = Convert.ToDateTime(reader.GetValue(3)).ToShortDateString();
                        info.Value[4] = reader.GetValue(4);
                        info.Value[5] = reader.GetString(5);
                        info.Value[6] = reader.GetValue(6);
                        info.Value[7] = Convert.ToDateTime(reader.GetValue(7)).ToShortDateString();
                        list.Add(info);
                    }
                Word.Table oTable;
                Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                oTable = oDoc.Tables.Add(wrdRng, list.Count + 2, 8, ref oMissing, ref oMissing);
                oTable.Range.ParagraphFormat.SpaceAfter = 6;
                oTable.Range.Font.Bold = 0;
                oTable.Range.Font.Size = 8;
                oTable.Cell(1, 1).Range.Text = "Зачетная книжка";
                oTable.Cell(1, 2).Range.Text = "Место";
                oTable.Cell(1, 3).Range.Text = "Дата начала";
                oTable.Cell(1, 4).Range.Text = "Дата окончания";
                oTable.Cell(1, 5).Range.Text = "Рук-ль от вуза";
                oTable.Cell(1, 6).Range.Text = "Рук-ть от пред.";
                oTable.Cell(1, 7).Range.Text = "Оценка";
                oTable.Cell(1, 8).Range.Text = "Дата выст-я";
                for (int i = 0; i < 9; i++)
                {
                    oTable.Cell(1, i).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleInset;
                }
                double dR = 0;
                for (int i = 2; i <= list.Count + 1; i++)
                    for (int j = 1; j <= 8; j++)
                    {
                        oTable.Cell(i, j).Range.Text = list[i - 2].Value[j - 1].ToString();
                        oTable.Cell(i, j).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleInset;
                        if (j == 7)
                        {
                            switch (list[i - 2].Value[j - 1].ToString())
                            {
                                case "Отл": dR += 5; break;
                                case "Хор": dR += 4; break;
                                case "Удв": dR += 3; break;
                                case "Неуд": dR += 2; break;
                            }
                        }
                    }

                oTable.Cell(list.Count + 2, 6).Range.Text = "Средняя оценка:";
                oTable.Cell(list.Count + 2, 6).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleInset;

                oTable.Cell(list.Count + 2, 7).Range.Text = (dR / list.Count).ToString("0.000");
                oTable.Cell(list.Count + 2, 7).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleInset;
                reader.Dispose();
                reader.Close();
                command.Dispose();
                connection.Close();
                connection.Dispose();
            }
            catch (Exception) { }
        }

        public void ReportOnYear(string Group, string Year)
        {
            object oMissing = Missing.Value;
            object oEndOfDoc = "\\endofdoc";
            Word.Application oWord = new Word.Application();
            Word.Document oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
            ref oMissing, ref oMissing);
            oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
            oWord.Visible = true;
            try
            {
                object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                OleDbConnection connection = new OleDbConnection(@"Provider =  Microsoft.Jet.OLEDB.4.0; Data Source = " + Program.PracticeBD);
                List<infoStd> list = new List<infoStd>();
                connection.Open();
                infoStd info;
                Word.Paragraph oPara1;
                oPara1 = oDoc.Content.Paragraphs.Add(ref oMissing);
                if (Year == "Все")
                {
                    oPara1.Range.Text = String.Format("Отчет по группе {0}", Group);
                    command = new OleDbCommand($"SELECT Практика.[Зачетная книжка], Практика.[Название практики], Практика.[Место прохождения], Практика.[Дата начала], Практика.[Дата окончания], Преподаватель.ФИО, Практика.[Руководитель от предприятия], Практика.Оценка, Практика.[Дата выставления] FROM Студент INNER JOIN (Преподаватель INNER JOIN Практика ON Преподаватель.ID = Практика.[Руководитель от вуза]) ON Студент.[Номер зачетки] = Практика.[Зачетная книжка] WHERE Студент.[Код учебной группы]= '{Group}'", connection);
                }
                else
                {
                    oPara1.Range.Text = String.Format("Отчет по группе {0} в {1} г.", Group, Year);
                    command = new OleDbCommand($"SELECT Практика.[Зачетная книжка], Практика.[Название практики], Практика.[Место прохождения], Практика.[Дата начала], Практика.[Дата окончания], Преподаватель.ФИО, Практика.[Руководитель от предприятия], Практика.Оценка, Практика.[Дата выставления] FROM Студент INNER JOIN (Преподаватель INNER JOIN Практика ON Преподаватель.ID = Практика.[Руководитель от вуза]) ON Студент.[Номер зачетки] = Практика.[Зачетная книжка] WHERE YEAR(Практика.[Дата начала])= '{Year}' AND Студент.[Код учебной группы]= '{Group}'", connection);
                }
                oPara1.Range.Font.Bold = 1;
                oPara1.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                oPara1.Format.SpaceAfter = 24;
                oPara1.Range.InsertParagraphAfter();
                reader = command.ExecuteReader();
                if (reader.HasRows)
                    while (reader.Read())
                    {
                        info = new infoStd(9);
                        info.Value[0] = reader.GetString(0);
                        info.Value[1] = reader.GetString(1);
                        info.Value[2] = reader.GetString(2);
                        info.Value[3] = Convert.ToDateTime(reader.GetValue(3)).ToShortDateString();
                        info.Value[4] = Convert.ToDateTime(reader.GetValue(4)).ToShortDateString();
                        info.Value[5] = reader.GetValue(5);
                        info.Value[6] = reader.GetString(6);
                        info.Value[7] = reader.GetValue(7);
                        info.Value[8] = Convert.ToDateTime(reader.GetValue(8)).ToShortDateString();
                        list.Add(info);
                    }
                Word.Table oTable;
                Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                oTable = oDoc.Tables.Add(wrdRng, list.Count + 2, 9, ref oMissing, ref oMissing);
                oTable.Range.ParagraphFormat.SpaceAfter = 6;
                oTable.Range.Font.Bold = 0;
                oTable.Range.Font.Size = 8;
                oTable.Cell(1, 1).Range.Text = "Зачетная книжка";
                oTable.Cell(1, 2).Range.Text = "Название практики";
                oTable.Cell(1, 3).Range.Text = "Место";
                oTable.Cell(1, 4).Range.Text = "Дата начала";
                oTable.Cell(1, 5).Range.Text = "Дата окончания";
                oTable.Cell(1, 6).Range.Text = "Рук-ль от вуза";
                oTable.Cell(1, 7).Range.Text = "Рук-ть от пред.";
                oTable.Cell(1, 8).Range.Text = "Оценка";
                oTable.Cell(1, 9).Range.Text = "Дата выст-я";
                for (int i = 0; i < 10; i++)
                {
                    oTable.Cell(1, i).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleInset;
                }
                double dR = 0;
                for (int i = 2; i <= list.Count + 1; i++)
                    for (int j = 1; j <= 9; j++)
                    {
                        oTable.Cell(i, j).Range.Text = list[i - 2].Value[j - 1].ToString();
                        oTable.Cell(i, j).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleInset;
                        if (j == 8)
                        {
                            switch (list[i - 2].Value[j - 1].ToString())
                            {
                                case "Отл": dR += 5; break;
                                case "Хор": dR += 4; break;
                                case "Удв": dR += 3; break;
                                case "Неуд": dR += 2; break;
                            }
                        }
                    }

                oTable.Cell(list.Count + 2, 7).Range.Text = "Средняя оценка:";
                oTable.Cell(list.Count + 2, 7).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleInset;

                oTable.Cell(list.Count + 2, 8).Range.Text = (dR / list.Count).ToString("0.000");
                oTable.Cell(list.Count + 2, 8).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleInset;
                reader.Dispose();
                reader.Close();
                command.Dispose();
                connection.Close();
                connection.Dispose();
            }
            catch (Exception) { }
        }

        public void ReportOnTeacher(string Name, string Year)
        {
            object oMissing = Missing.Value;
            object oEndOfDoc = "\\endofdoc";
            Word.Application oWord = new Word.Application();
            Word.Document oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
            ref oMissing, ref oMissing);
            oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
            oWord.Visible = true;
            try
            {
                object oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                OleDbConnection connection = new OleDbConnection(@"Provider =  Microsoft.Jet.OLEDB.4.0; Data Source = " + Program.PracticeBD);
                List<infoStd> list = new List<infoStd>();
                connection.Open();
                infoStd info;
                Word.Paragraph oPara1;
                oPara1 = oDoc.Content.Paragraphs.Add(ref oMissing);
                if (Year == "Все")
                {
                    oPara1.Range.Text = String.Format("Отчет по руководителю от вуз - {0}", Name);
                    command = new OleDbCommand($"SELECT Практика.[Зачетная книжка], Практика.[Название практики], Практика.[Место прохождения], Практика.[Дата начала], Практика.[Дата окончания], Практика.[Руководитель от предприятия], Практика.Оценка, Практика.[Дата выставления], Практика.[Дата начала] FROM Студент INNER JOIN (Преподаватель INNER JOIN Практика ON Преподаватель.ID = Практика.[Руководитель от вуза]) ON Студент.[Номер зачетки] = Практика.[Зачетная книжка] WHERE Преподаватель.ФИО= '{Name}'", connection);
                }
                else
                {
                    oPara1.Range.Text = String.Format("Отчет по руководителю от вуза {0} за {1} г.", Name, Year);
                    command = new OleDbCommand($"SELECT Практика.[Зачетная книжка], Практика.[Название практики], Практика.[Место прохождения], Практика.[Дата начала], Практика.[Дата окончания], Практика.[Руководитель от предприятия], Практика.Оценка, Практика.[Дата выставления], Практика.[Дата начала] FROM Студент INNER JOIN (Преподаватель INNER JOIN Практика ON Преподаватель.ID = Практика.[Руководитель от вуза]) ON Студент.[Номер зачетки] = Практика.[Зачетная книжка] WHERE Преподаватель.ФИО= '{Name}' AND YEAR(Практика.[Дата начала])= '{Year}'", connection);
                }
                oPara1.Range.Font.Bold = 1;
                oPara1.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                oPara1.Format.SpaceAfter = 24;
                oPara1.Range.InsertParagraphAfter();
                reader = command.ExecuteReader();
                if (reader.HasRows)
                    while (reader.Read())
                    {
                        info = new infoStd(8);
                        info.Value[0] = reader.GetString(0);
                        info.Value[1] = reader.GetString(1);
                        info.Value[2] = reader.GetString(2);
                        info.Value[3] = Convert.ToDateTime(reader.GetValue(3)).ToShortDateString();
                        info.Value[4] = Convert.ToDateTime(reader.GetValue(4)).ToShortDateString();
                        info.Value[5] = reader.GetValue(5);
                        info.Value[6] = reader.GetValue(6);
                        info.Value[7] = Convert.ToDateTime(reader.GetValue(7)).ToShortDateString();
                        list.Add(info);
                    }
                Word.Table oTable;
                Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                oTable = oDoc.Tables.Add(wrdRng, list.Count + 2, 9, ref oMissing, ref oMissing);
                oTable.Range.ParagraphFormat.SpaceAfter = 6;
                oTable.Range.Font.Bold = 0;
                oTable.Range.Font.Size = 8;
                oTable.Cell(1, 1).Range.Text = "Зачетная книжка";
                oTable.Cell(1, 2).Range.Text = "Название практики";
                oTable.Cell(1, 3).Range.Text = "Место";
                oTable.Cell(1, 4).Range.Text = "Дата начала";
                oTable.Cell(1, 5).Range.Text = "Дата окончания";
                oTable.Cell(1, 6).Range.Text = "Рук-ть от пред.";
                oTable.Cell(1, 7).Range.Text = "Оценка";
                oTable.Cell(1, 8).Range.Text = "Дата выст-я";
                for (int i = 0; i < 9; i++)
                {
                    oTable.Cell(1, i).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleInset;
                }
                double dR = 0;
                for (int i = 2; i <= list.Count + 1; i++)
                    for (int j = 1; j <= 8; j++)
                    {
                        oTable.Cell(i, j).Range.Text = list[i - 2].Value[j - 1].ToString();
                        oTable.Cell(i, j).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleInset;
                        if (j == 7)
                        {
                            switch (list[i - 2].Value[j - 1].ToString())
                            {
                                case "Отл": dR += 5; break;
                                case "Хор": dR += 4; break;
                                case "Удв": dR += 3; break;
                                case "Неуд": dR += 2; break;
                            }
                        }
                    }

                oTable.Cell(list.Count + 2, 6).Range.Text = "Средняя оценка:";
                oTable.Cell(list.Count + 2, 6).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleInset;

                oTable.Cell(list.Count + 2, 7).Range.Text = (dR / list.Count).ToString("0.000");
                oTable.Cell(list.Count + 2, 7).Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleInset;
                reader.Dispose();
                reader.Close();
                command.Dispose();
                connection.Close();
                connection.Dispose();
            }
            catch (Exception) { }
        }
        #endregion 
    }
}