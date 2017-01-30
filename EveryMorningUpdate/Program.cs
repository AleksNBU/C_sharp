using System;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Linq;
namespace NS
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void MadeProc(string a, string b)
            {
                ProcessStartInfo psi = new ProcessStartInfo(a, b);
                    Process Proc = new Process();
                    Proc = Process.Start(psi);
                    Proc.WaitForExit();
            }
                 
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            string netDPTLIST = @"w:\nsi_dbf\DPTLIST.arj";
            string netRCUKRU = @"w:\nsi_dbf\RCUKRU.arj";
            string netCOMPANY = @"w:\nsi_dbf\COMPANY.arj";
            string netRC_BNK = @"w:\nsi_dbf\rc_bnk\RC_BNK.dbf";
            string netFRNBANKS = @"w:\nsi_dbf\rc_bnk\FRNBANKS.dbf";
          
            string localDPTLIST = @"D:\NEW_STAT\spr_form\DPTLIST.arj";
            string localRCUKRU = @"D:\NEW_STAT\spr_form\RCUKRU.arj";
            string localCOMPANY = @"D:\NEW_STAT\spr_form\COMPANY.arj";
            string localRC_BNK = @"D:\NEW_STAT\spr_form\RC_BNK.dbf";
            string localFRNBANKS = @"D:\NEW_STAT\spr_form\FRNBANKS.dbf";

            string obm_rcu = @"Q:\PROCARRY\NEWSTAT\IN\OBM_RCU.BAT";
            string str1 = "";
            string str2 = "";

            if (!System.IO.File.Exists(netDPTLIST) || !System.IO.File.Exists(netRCUKRU) || !System.IO.File.Exists(netRC_BNK) || !System.IO.File.Exists(netFRNBANKS) || !System.IO.File.Exists(netCOMPANY))
            {
                MessageBox.Show("Проблемы с сетью \n\nПроверьте путь \n\nftp://pluto.bank.gov.ua/hosting/nsi_dbf/", "Внимание!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Environment.Exit(0);
            }
            if (!System.IO.File.Exists(obm_rcu))
            {
                MessageBox.Show("Проблемы с сетевым диском Q\n\n  Проверьте доступ к ящику\n\n             NEWSTAT", "Внимание!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Environment.Exit(0);
            }

            foreach (string file in Directory.GetFiles(@"D:\NEW_STAT\SPR_FORM", "*.*", SearchOption.AllDirectories)
                .Where(s => s.ToUpper().Contains("COMPANY") || s.ToUpper().Contains("DPTLIST") || s.ToUpper().Contains("RCUKRU") || s.ToUpper().Contains("FRNBANKS") || s.ToUpper().Contains("RC_BNK")))
            {
                File.SetAttributes(file, FileAttributes.Normal);
                File.Delete(file);
            }

            File.Copy(netDPTLIST, localDPTLIST, true);
            File.Copy(netRCUKRU, localRCUKRU, true);
            File.Copy(netCOMPANY, localCOMPANY, true);

            Process.Start(@"D:\ARC\arj.exe", @"e -y d:\new_stat\spr_form\DPTLIST.ARJ d:\new_stat\spr_form\");
            Process.Start(@"D:\ARC\arj.exe", @"e -y d:\new_stat\spr_form\RCUKRU.ARJ d:\new_stat\spr_form\");
            Process.Start(@"D:\ARC\arj.exe", @"e -y d:\new_stat\spr_form\COMPANY.ARJ d:\new_stat\spr_form\");

            FileInfo fileInfonetRC_BNK = new FileInfo(netRC_BNK);
            FileInfo fileInfolocalRC_BNK = new FileInfo(@"D:\NEW_STAT\kod_form\RC_BNK.DBF");
            FileInfo fileInfonetFRNBANKS = new FileInfo(netFRNBANKS);
            FileInfo fileInfolocalFRNBANKS = new FileInfo(@"D:\NEW_STAT\kod_form\FRNBANKS.DBF");

            if (fileInfonetRC_BNK.LastWriteTime > fileInfolocalRC_BNK.LastWriteTime)
            {
                str1 = "Новый RC_BNK.DBF\n";
                File.Copy(netRC_BNK, localRC_BNK, true);
                File.SetAttributes(localRC_BNK, FileAttributes.Normal);
                File.Copy(localRC_BNK, @"D:\NEW_STAT\kod_form\RC_BNK.DBF", true);
            }
            if (fileInfonetFRNBANKS.LastWriteTime > fileInfolocalFRNBANKS.LastWriteTime)
            {
                str2 = "Новый FRNBANKS.DBF\n";
                File.Copy(netFRNBANKS, localFRNBANKS, true);
                File.SetAttributes(localFRNBANKS, FileAttributes.Normal);
                File.Copy(localFRNBANKS, @"D:\NEW_STAT\kod_form\FRNBANKS.DBF", true);
            }
            Process.Start(obm_rcu);
            MadeProc("RAR.exe", @"a -y d:\new_stat\spr_form\DPTLIST.RAR d:\new_stat\spr_form\DPTLIST.DBF");
            MadeProc(@"q:\nbu_post\nbumail\bin\tomail.exe", @"-fd:\new_stat\spr_form\dptlist.rar -nnewstat -snewstat -h#1#");
            MadeProc(@"q:\nbu_post\nbumail\bin\tomail.exe", @"-fd:\new_stat\spr_form\dptlist.rar -nnewstat -snewstat -h#2#");

            File.Copy(@"d:\new_stat\spr_form\COMPANY.DBF", @"D:\NEW_STAT\kod_form\COMPANY.DBF", true);
            File.Delete(@"d:\new_stat\spr_form\DPTLIST.RAR");

            MessageBox.Show("Файлы скачены и распакованы\nDPTLIST разослан банкам\n\n" + str1 + str2 + "\nОбновите справочники в Стате!\nСделайте обновление ключей!",
                "Обновление СтатОтчетности", MessageBoxButtons.OK, MessageBoxIcon.Information);
            Application.Exit();
        }
    }
}