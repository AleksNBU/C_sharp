using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Drawing;

namespace PeopleNBU
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            using (var mutex = new System.Threading.Mutex(false, "cyberforum.ru example"))
   {
       if (mutex.WaitOne(TimeSpan.FromSeconds(3)))
       {
           Application.Run(new Splash());
           Application.Run(new Form1());
       }
       else
           MessageBox.Show("Один экземпляр приложения уже запущен!");
   }
        }
    }
    static class GlobalData
    {
        public static int id { get; set; }
        public static string Tab { get; set; }
        public static string F { get; set; }
        public static string I { get; set; }
        public static string O { get; set; }
        public static DateTime Bithday { get; set; }
        public static string Otdel { get; set; }
        public static string Dolzhn { get; set; }
        public static string Tel { get; set; }
        public static string Comp { get; set; }
        public static string Email { get; set; }
        public static string Adress { get; set; }
        public static string BIRTHPLACE { get; set; }
        public static string Passport { get; set; }
        public static string IPN { get; set; }
        public static string[] arrayTabNumber;
        public static Image pictureBox { get; set; }
        public static string querySQL { get; set; }
        public static int prizn { get; set; }
    }
    static class GlobalData1
    {
        public static Image pictureBox { get; set; }
    }
}
