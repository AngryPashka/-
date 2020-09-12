using System.Media;
namespace Учет_прохождения_практики_студентами
{
    #region Класс воспроизведения звуков
    public class Music
    {
        public static void PlayClick()
        {
            SoundPlayer Cl = new SoundPlayer(Properties.Resources.Click);
            Cl.Load();
            Cl.Play();
        }

        public static void PlayError()
        {
            SoundPlayer Cl = new SoundPlayer(Properties.Resources.Ошибка);
            Cl.Load();
            Cl.Play();
        }

    }
    #endregion
}
