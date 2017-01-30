using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

namespace Media
{
    public class Player
    {

        private String Pcommand;
        private bool isOpen;

        [DllImport("winmm.dll")]
        //[DllImport("Winmm.dll", CharSet = CharSet.Auto)]
        private static extern long mciSendString(string strCommand, StringBuilder strReturn, int iReturnLength, IntPtr hwndCallback);
        private static Player player;
        private int masterVolumn;

        public Player()
        {
            this.MasterVolume = 10 * 50;
        }

        public static Player GetPlayer()
        {
            if (player == null)
                player = new Player();
            return player;
        }

        public void Close()
        {
            Pcommand = "close MediaFile";
            mciSendString(Pcommand, null, 0, IntPtr.Zero);
            isOpen = false;
        }

        public void Open(string sFileName)
        {
            Pcommand = "open \"" + sFileName + "\" type mpegvideo alias MediaFile";
            mciSendString(Pcommand, null, 0, IntPtr.Zero);
            isOpen = true;
        }


        /// <summary>
        /// Воспроизведение уже открытого файла по кругу или нет
        /// </summary>
        public void Play(bool loop)
        {
            if (isOpen)
            {
                Pcommand = "play MediaFile";
                if (loop)
                    Pcommand += " REPEAT";
                mciSendString(Pcommand, null, 0, IntPtr.Zero);
                this.MasterVolume = this.MasterVolume;
            }
        }

        /// <summary>
        /// Воспроизведение указаного файла
        /// </summary>
        public void Play(string FileName)
        {
            if (isOpen == true)
            {
                Close();
            }
            Open(FileName);
            Play(true);
        }

        /// <summary>
        /// Установка паузы
        /// </summary>
        public void Pause()
        {
            Pcommand = "pause MediaFile";
            mciSendString(Pcommand, null, 0, IntPtr.Zero);
        }

        /// <summary>
        /// Получение текущего статуса
        /// </summary>
        public String Status()
        {
            int i = 128;
            System.Text.StringBuilder stringBuilder = new System.Text.StringBuilder(i);
            mciSendString("status MediaFile mode", stringBuilder, i, IntPtr.Zero);
            return stringBuilder.ToString();
        }

        /// <summary>
        /// Получение/Установка общей громкости воспроизведения
        /// </summary>
        public int MasterVolume
        {
            get
            {
                return masterVolumn;
            }
            set
            {
                mciSendString(string.Concat("setaudio MediaFile volume to ", value), null, 0, IntPtr.Zero);
                masterVolumn = value;
            }
        }

        /// <summary>
        /// Проверка установленна ли пауза
        /// </summary>
        public bool IsPaused()
        {
            return Pcommand == "pause MediaFile";
        }

        /// <summary>
        /// Проверка происходит ли проигрывание
        /// </summary>
        public bool IsPlaying()
        {
            return Status() == "playing";
        }

        /// <summary>
        /// Проверка Открыт ли какой либо файл
        /// </summary>
        public bool IsOpen()
        {
            return isOpen;
        }

    }

}
