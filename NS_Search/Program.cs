using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;

namespace NS_Obnovlenie
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
            string dirNewstat = @"Q:\PROCARRY\NEWSTAT\IN\";
            string fileNewstat = @"D:\NEW_STAT\SPR_FORM\!!!\Update\Dateoflastupdate.txt";
            if (!Directory.Exists(dirNewstat))
            {
                MessageBox.Show("Проблемы с сетью \n\nПроверьте путь к ящику\n\n      NEWSTAT", "Внимание!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Environment.Exit(0);
            }

            string[] array1 = Directory.GetFiles(dirNewstat, "*.arj");
            string[] array2 = Directory.GetFiles(dirNewstat, "NBU24SPR.*");
            string[]array3 = array1.Concat(array2).ToArray();
            
            if (array3.Length != 0)
            {
                string[][] a = new string[array3.Length][];
                int k = 0;
                string str = " ";
                foreach (string element in array3)
                {
                    a[k] = new string[] { Path.GetFileName(element), File.GetLastWriteTime(element).ToString() };
                    k++;
                }
                
                Array.Sort(a, (a1, a2) => { return DateTime.Parse(a2[1]).CompareTo(DateTime.Parse(a1[1])); });
                
                string readText;
                if (File.Exists(fileNewstat))
                {
                    FileInfo f = new FileInfo(fileNewstat);
                    var s1 = f.Length;
                    if (s1 > 0)
                        readText = File.ReadAllText(fileNewstat);
                    else
                        readText = "01.01.2000 0:00:00";
                }
                else
                {
                    File.Create(fileNewstat).Close();
                    readText = "01.01.2000 0:00:00";
                }
                bool x = true;
                int y = 0;
                str = "";
                for (int i = 0; (i < a.Length) & x; i++)
                {
                    if (DateTime.Parse(a[i][1]) <= DateTime.Parse(readText))
                    {
                        x = false;
                    }
                    else
                    {
                        y++;
                        str +=  "                 " + a[i][0] + "\n\n";
                        x = true;
                        if (i == 0)
                        {
                            File.WriteAllText(fileNewstat, a[i][1]);
                        }
                    }
                }
                if (y != 0)
                {
                    MessageBox.Show("!!! ОБНОВИТЕ СТАТОТЧЕТНОСТЬ !!!\n\n  ----- Есть новые обновления -----\n        ----   в количестве (" + y + ")  ----\n\n" + str,
                    "Поиск Обновлений. Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                Application.Exit();
            }
            else
           {
                MessageBox.Show("В ящике NEWSTAT архивных файлов нет.", "Поиск Обновлений.", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Application.Exit();
            }
        }
    }
}
