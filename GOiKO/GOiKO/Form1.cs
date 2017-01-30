using System;
using System.Threading;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using System.ComponentModel;
using System.Text.RegularExpressions;
namespace GOiKO
{
    public partial class Form1 : Form
    {
        string FilesDBF, FilesAccess;
        string DAT = "";
        int i = 0;
        int y = 0;
        int z = 0;
        int kkk = 0;
        int ttt = 0;
        string[] Mdate;
        OleDbConnection cn2, cn1;
        DataTable table1, table2;
        OleDbDataAdapter adapter1, adapter2;
        OleDbDataReader reader2, reader;
        OleDbCommand cmd1, cmd;
        public Form1()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (backgroundWorker1.IsBusy != true)
            {
                    string strDBConnect1 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + @Path.GetDirectoryName(FilesDBF) + ";Extended Properties=dBASE IV";
                    cn1 = new OleDbConnection(strDBConnect1);
                    adapter1 = new OleDbDataAdapter("SELECT * FROM " + Path.GetFileNameWithoutExtension(FilesDBF), cn1);
                    table1 = new DataTable();
                    try
                    {
                        cn1.Open();
                        adapter1.Fill(table1);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Помилка відкриття файлу 748.dbf! " + ex.Message, "Увага, помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    string strDBConnect2 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " + @Path.GetFullPath(FilesAccess) + ";Persist Security Info=False";
                     cn2 = new OleDbConnection(strDBConnect2);
                     adapter2 = new OleDbDataAdapter("SELECT * FROM [Касові обороти]", cn2);
                     table2 = new DataTable();
                    try
                    {
                        cn2.Open();
                        adapter2.Fill(table2);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Помилка відкриття файлу 748.accdb! " + ex.Message, "Увага, помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                        if ((table1.Columns[0].ColumnName == "DT") & (table1.Columns[1].ColumnName == "NKB") &
                            (table1.Columns[2].ColumnName == "KU") & (table1.Columns[3].ColumnName == "TXT") &
                            (table1.Columns[4].ColumnName == "V01"))
                        {

                        labelMessage.Text = "Початок процесу";
                        labelMessage.Visible = true;
                        progressBar1.Visible = true;
                        // Start the asynchronous operation.
                        backgroundWorker1.RunWorkerAsync();
                    //}
//==============================================================================================================
                        }
                        else
                        {
                            MessageBox.Show("Структура файлу 748.dbf не відповідає 748 формі Статзвітності!", "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            cn1.Close();
                            cn2.Close();
                            textBox1.Clear();
                            textBox2.Clear();
                            button2.Enabled = false;
                            progressBar1.Visible = false;
                            button3.Enabled = false;
                            button1.Focus();
                        }
            }
        
        }

        private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;

            cmd1 = cn2.CreateCommand();
            cmd1.CommandText = "SELECT Дата FROM [Касові обороти] GROUP BY Дата";
            reader2 = cmd1.ExecuteReader();
            String format = "yyyy-MM-dd";
            while (reader2.Read())
            {
                kkk++;
            }
            reader2.Close();
            reader2 = cmd1.ExecuteReader();
            
            if (kkk > 0)
            {
                DAT = "WHERE mt.dt not in (";
                Mdate = new string[kkk];
                while (reader2.Read())
                {
                   Mdate[ttt++] = "#" + reader2.GetDateTime(0).ToString(format) + "#";
                 }
                foreach (string qqqq in Mdate)
                {
                   DAT += qqqq + ",";
                }
               DAT = DAT.Remove(DAT.Length - 1, 1) + ")";
            }

            cmd = cn1.CreateCommand();
            cmd.CommandText = "SELECT mt.DT AS Дата, mt.NKB AS [Код банку], mt.KU AS [Код області], " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='02' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 2, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='05' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 5, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='12' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 12, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='14' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 14, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='16' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 16, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='17' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 17, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='29' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 29, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='30' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 30, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='31' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 31, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='32' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 32, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='34' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 34, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='35' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 35, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='36' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 36, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='37' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 37, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='38' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 38, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='39' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 39, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='40' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 40, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='45' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 45, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='46' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 46, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='50' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 50, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='53' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 53, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='55' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 55, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='56' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 56, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='58' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 58, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='59' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 59, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='60' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 60, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='61' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 61, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='66' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 66, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='69' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 69, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='70' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 70, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='71' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 71, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='72' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 72, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='73' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 73, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='84' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 84, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='86' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 86, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='87' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 87, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='88' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 88, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='93' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 93, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='94' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 94, " +
            "(select [748].[V01] from 748 where left(ltrim([748].[TXT]),2)='95' and dt=mt.dt and nkb=mt.nkb and ku=mt.ku) AS 95 " +
            "FROM 748 AS mt " +
            DAT +
            "GROUP BY mt.DT, mt.NKB, mt.KU";

            reader = cmd.ExecuteReader();

            while (reader.Read())
            {
                i++;
            }
            
            reader.Close();
            reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                     string sql3 = "INSERT INTO [Касові обороти] (Дата, [Код банку], [Код області], " +
                    "2, 5, 12, 14, 16, 17, 29, 30, 31, 32, 34, 35, 36, 37, 38, 39, 40, 45, 46, 50, 53, 55, 56, 58, 59, 60, 61, 66, 69, 70, 71, 72, 73, 84, 86, 87, 88, 93, 94, 95 )" +
                    "VALUES (@qqq1, @qqq2, @qqq3, @qqq4, @qqq5, @qqq6, @qqq7, @qqq8, @qqq9, @qqq10, @qqq11, @qqq12, @qqq13, @qqq14, @qqq15, @qqq16, @qqq17, @qqq18, @qqq19, @qqq20," +
                    "@qqq21, @qqq22, @qqq23, @qqq24, @qqq25, @qqq26, @qqq27, @qqq28, @qqq29, @qqq30, @qqq31, @qqq32, @qqq33, @qqq34, @qqq35, @qqq36, @qqq37, @qqq38, @qqq39, @qqq40," +
                    "@qqq41, @qqq42, @qqq43)";
                    string qqq4, qqq5, qqq6, qqq7, qqq8, qqq9, qqq10, qqq11, qqq12, qqq13, qqq14, qqq15, qqq16, qqq17, qqq18, qqq19, qqq20,
                    qqq21, qqq22, qqq23, qqq24, qqq25, qqq26, qqq27, qqq28, qqq29, qqq30, qqq31, qqq32, qqq33, qqq34, qqq35, qqq36, qqq37, qqq38, qqq39, qqq40,
                    qqq41, qqq42, qqq43;

                    if (reader.IsDBNull(3))
                        qqq4 = "0";
                    else qqq4 = reader.GetString(3);

                    if (reader.IsDBNull(4))
                        qqq5 = "0";
                    else qqq5 = reader.GetString(4);

                    if (reader.IsDBNull(5))
                        qqq6 = "0";
                    else qqq6 = reader.GetString(5);

                    if (reader.IsDBNull(6))
                        qqq7 = "0";
                    else qqq7 = reader.GetString(6);

                    if (reader.IsDBNull(7))
                        qqq8 = "0";
                    else qqq8 = reader.GetString(7);

                    if (reader.IsDBNull(8))
                        qqq9 = "0";
                    else qqq9 = reader.GetString(8);

                    if (reader.IsDBNull(9))
                        qqq10 = "0";
                    else qqq10 = reader.GetString(9);

                    if (reader.IsDBNull(10))
                        qqq11 = "0";
                    else qqq11 = reader.GetString(10);

                    if (reader.IsDBNull(11))
                        qqq12 = "0";
                    else qqq12 = reader.GetString(11);

                    if (reader.IsDBNull(12))
                        qqq13 = "0";
                    else qqq13 = reader.GetString(12);

                    if (reader.IsDBNull(13))
                        qqq14 = "0";
                    else qqq14 = reader.GetString(13);

                    if (reader.IsDBNull(14))
                        qqq15 = "0";
                    else qqq15 = reader.GetString(14);

                    if (reader.IsDBNull(15))
                        qqq16 = "0";
                    else qqq16 = reader.GetString(15);

                    if (reader.IsDBNull(16))
                        qqq17 = "0";
                    else qqq17 = reader.GetString(16);

                    if (reader.IsDBNull(17))
                        qqq18 = "0";
                    else qqq18 = reader.GetString(17);

                    if (reader.IsDBNull(18))
                        qqq19 = "0";
                    else qqq19 = reader.GetString(18);

                    if (reader.IsDBNull(19))
                        qqq20 = "0";
                    else qqq20 = reader.GetString(19);

                    if (reader.IsDBNull(20))
                        qqq21 = "0";
                    else qqq21 = reader.GetString(20);

                    if (reader.IsDBNull(21))
                        qqq22 = "0";
                    else qqq22 = reader.GetString(21);

                    if (reader.IsDBNull(22))
                        qqq23 = "0";
                    else qqq23 = reader.GetString(22);

                    if (reader.IsDBNull(23))
                        qqq24 = "0";
                    else qqq24 = reader.GetString(23);

                    if (reader.IsDBNull(24))
                        qqq25 = "0";
                    else qqq25 = reader.GetString(24);

                    if (reader.IsDBNull(25))
                        qqq26 = "0";
                    else qqq26 = reader.GetString(25);

                    if (reader.IsDBNull(26))
                        qqq27 = "0";
                    else qqq27 = reader.GetString(26);

                    if (reader.IsDBNull(27))
                        qqq28 = "0";
                    else qqq28 = reader.GetString(27);

                    if (reader.IsDBNull(28))
                        qqq29 = "0";
                    else qqq29 = reader.GetString(28);

                    if (reader.IsDBNull(29))
                        qqq30 = "0";
                    else qqq30 = reader.GetString(29);

                    if (reader.IsDBNull(30))
                        qqq31 = "0";
                    else qqq31 = reader.GetString(30);

                    if (reader.IsDBNull(31))
                        qqq32 = "0";
                    else qqq32 = reader.GetString(31);

                    if (reader.IsDBNull(32))
                        qqq33 = "0";
                    else qqq33 = reader.GetString(32);

                    if (reader.IsDBNull(33))
                        qqq34 = "0";
                    else qqq34 = reader.GetString(33);

                    if (reader.IsDBNull(34))
                        qqq35 = "0";
                    else qqq35 = reader.GetString(34);

                    if (reader.IsDBNull(35))
                        qqq36 = "0";
                    else qqq36 = reader.GetString(35);

                    if (reader.IsDBNull(36))
                        qqq37 = "0";
                    else qqq37 = reader.GetString(36);

                    if (reader.IsDBNull(37))
                        qqq38 = "0";
                    else qqq38 = reader.GetString(37);

                    if (reader.IsDBNull(38))
                        qqq39 = "0";
                    else qqq39 = reader.GetString(38);

                    if (reader.IsDBNull(39))
                        qqq40 = "0";
                    else qqq40 = reader.GetString(39);

                    if (reader.IsDBNull(40))
                        qqq41 = "0";
                    else qqq41 = reader.GetString(40);

                    if (reader.IsDBNull(41))
                        qqq42 = "0";
                    else qqq42 = reader.GetString(41);

                    if (reader.IsDBNull(42))
                        qqq43 = "0";
                    else qqq43 = reader.GetString(42);

                    using (OleDbCommand cmd2 = new OleDbCommand(sql3, cn2))
                    {
                        cmd2.Parameters.AddWithValue("@qqq1", reader.GetDateTime(0).ToShortDateString());
                        cmd2.Parameters.AddWithValue("@qqq2", reader.GetString(1));
                        cmd2.Parameters.AddWithValue("@qqq3", reader.GetString(2));
                        cmd2.Parameters.AddWithValue("@qqq4", qqq4);
                        cmd2.Parameters.AddWithValue("@qqq5", qqq5);
                        cmd2.Parameters.AddWithValue("@qqq6", qqq6);
                        cmd2.Parameters.AddWithValue("@qqq7", qqq7);
                        cmd2.Parameters.AddWithValue("@qqq8", qqq8);
                        cmd2.Parameters.AddWithValue("@qqq9", qqq9);
                        cmd2.Parameters.AddWithValue("@qqq10", qqq10);
                        cmd2.Parameters.AddWithValue("@qqq11", qqq11);
                        cmd2.Parameters.AddWithValue("@qqq12", qqq12);
                        cmd2.Parameters.AddWithValue("@qqq13", qqq13);
                        cmd2.Parameters.AddWithValue("@qqq14", qqq14);
                        cmd2.Parameters.AddWithValue("@qqq15", qqq15);
                        cmd2.Parameters.AddWithValue("@qqq16", qqq16);
                        cmd2.Parameters.AddWithValue("@qqq17", qqq17);
                        cmd2.Parameters.AddWithValue("@qqq18", qqq18);
                        cmd2.Parameters.AddWithValue("@qqq19", qqq19);
                        cmd2.Parameters.AddWithValue("@qqq20", qqq20);
                        cmd2.Parameters.AddWithValue("@qqq21", qqq21);
                        cmd2.Parameters.AddWithValue("@qqq22", qqq22);
                        cmd2.Parameters.AddWithValue("@qqq23", qqq23);
                        cmd2.Parameters.AddWithValue("@qqq24", qqq24);
                        cmd2.Parameters.AddWithValue("@qqq25", qqq25);
                        cmd2.Parameters.AddWithValue("@qqq26", qqq26);
                        cmd2.Parameters.AddWithValue("@qqq27", qqq27);
                        cmd2.Parameters.AddWithValue("@qqq28", qqq28);
                        cmd2.Parameters.AddWithValue("@qqq29", qqq29);
                        cmd2.Parameters.AddWithValue("@qqq30", qqq30);
                        cmd2.Parameters.AddWithValue("@qqq31", qqq31);
                        cmd2.Parameters.AddWithValue("@qqq32", qqq32);
                        cmd2.Parameters.AddWithValue("@qqq33", qqq33);
                        cmd2.Parameters.AddWithValue("@qqq34", qqq34);
                        cmd2.Parameters.AddWithValue("@qqq35", qqq35);
                        cmd2.Parameters.AddWithValue("@qqq36", qqq36);
                        cmd2.Parameters.AddWithValue("@qqq37", qqq37);
                        cmd2.Parameters.AddWithValue("@qqq38", qqq38);
                        cmd2.Parameters.AddWithValue("@qqq39", qqq39);
                        cmd2.Parameters.AddWithValue("@qqq40", qqq40);
                        cmd2.Parameters.AddWithValue("@qqq41", qqq41);
                        cmd2.Parameters.AddWithValue("@qqq42", qqq42);
                        cmd2.Parameters.AddWithValue("@qqq43", qqq43);
                        cmd2.ExecuteNonQuery();
                    }
                z++;
                y++;
                worker.ReportProgress((y) * 100 / i);
                System.Threading.Thread.Sleep(0);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (backgroundWorker1.IsBusy != true)
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                openFileDialog1.Filter = "DBF file |748.dbf";
                openFileDialog1.Title = "Пошук DBF файла";
                openFileDialog1.Multiselect = false;

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        FilesDBF = openFileDialog1.FileName;
                        textBox1.Text = FilesDBF;
                        button2.Enabled = true;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Файл DBF не може бути прочитаний: " + ex.Message, "Увага, помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (backgroundWorker1.IsBusy != true)
            {
                OpenFileDialog openFileDialog2 = new OpenFileDialog();
                openFileDialog2.Filter = "Access file |748.accdb";
                openFileDialog2.Title = "Пошук Access файла";
                openFileDialog2.Multiselect = false;
                if (openFileDialog2.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        FilesAccess = openFileDialog2.FileName;
                        textBox2.Text = FilesAccess;
                        button3.Enabled = true;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Файл Access не може бути прочитаний: " + ex.Message, "Увага, помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void backgroundWorker1_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            labelMessage.Text = "Чекайте будь ласка, процес іде... " + e.ProgressPercentage.ToString() + "%";
            progressBar1.Value = e.ProgressPercentage;
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            table2.Clear();
            adapter2.Fill(table2);
            if (DAT.Length != 0)
            {
            DAT = "Інформація за наступні дати вже присутня у базі: ";
                reader2.Close();
                reader2 = cmd1.ExecuteReader();


                cmd = cn1.CreateCommand();
                cmd.CommandText = "SELECT DT FROM 748 GROUP BY DT";
                reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    while (reader2.Read())
                    {
                        if (reader2.GetDateTime(0) == reader.GetDateTime(0))
                        {
                            DAT += reader.GetDateTime(0).ToShortDateString() + ", ";
                            break;
                        }
                    }
                    reader2.Close();
                    reader2 = cmd1.ExecuteReader();
                }

                //DAT = DAT.Remove(DAT.Length - 1, 1).Substring(20);
                //var re = new Regex("#");
                //DAT = re.Replace(DAT, " ");
                //DAT = "Інформація за наступні дати присутня у базі: " + DAT; 
                DAT = DAT.Substring(0, DAT.Length - 2);
            }
            MessageBox.Show("Процес успішно завершений!\n\nКількість перенесених записів - " + z.ToString() +
            "\nЗагальна кількість записів у базі - " + table2.Rows.Count + "\n" + DAT, "OK!", MessageBoxButtons.OK, MessageBoxIcon.Information);  //MessageBox.Show(table2.Rows.Count.ToString());
            labelMessage.Visible = false;
            progressBar1.Value = 0;
            cn1.Close();
            cn2.Close();
            textBox1.Clear();
            textBox2.Clear();
            button2.Enabled = false;
            progressBar1.Visible = false;
            button3.Enabled = false;
            button1.Focus();
            i = y = z = kkk = ttt = 0;
            Mdate = null;
        }

    }

}