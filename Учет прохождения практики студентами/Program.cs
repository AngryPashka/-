using System;
using System.Windows.Forms;
using System.IO;

namespace Учет_прохождения_практики_студентами
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            #region Проверка файла и папок БД
            if (!Directory.Exists(@"C:\ProgramData\Учет"))
                Directory.CreateDirectory(@"C:\ProgramData\Учет");

            if (!File.Exists(UsersBD))
                File.WriteAllBytes(UsersBD, Properties.Resources.Users);

            if (!File.Exists(PracticeBD))
                File.WriteAllBytes(PracticeBD, Properties.Resources.Практика);
            #endregion
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new FormAuthentication());
        }

        public static string UsersBD = @"C:\ProgramData\Учет\Users.mdb";
        public static string PracticeBD = @"C:\ProgramData\Учет\Практика.mdb";
    }
}
