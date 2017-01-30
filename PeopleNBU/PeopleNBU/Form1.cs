using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using System.Data.OleDb;
using System.IO;
using System.Media;
using System.Text.RegularExpressions;
using System.Net.NetworkInformation;
using System.Diagnostics;

namespace PeopleNBU
{
    public partial class Form1 : Form
    {
        private string myAudio = Path.GetTempPath() + "sound.wma";
        private string pathDB = @"d:\MyDocuments\Табельные номера\NBU.accdb";
        private string strDBConnect;
        private string strSQL = "Select * from Table1";
        System.Reflection.Assembly a = System.Reflection.Assembly.GetEntryAssembly();
        DataSet dSet;
        private string ip;
        OleDbConnection cn;
        OleDbDataAdapter ad;
        OleDbCommand cmd;
        OleDbParameter imageParametr;
        Int32 maxid;
        private bool p3 = true;
        private bool flag = false;
        int v;
        int _page;
         double allPage;
        Font _font = new Font("Segoe UI", 14);

        public Form1()
        {
            InitializeComponent();

            GlobalData.prizn = 0;
            sound(myAudio);
            string baseDir = System.IO.Path.GetDirectoryName(a.Location);
            if (File.Exists(pathDB))
            {
                //strDBConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=d:\\MyDocuments\\Табельные номера\\NBU.accdb;Persist Security Info=False";

                //Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\myFolder\myAccess2007file.accdb;Jet OLEDB:Database Password=MyDbPassword; 



                strDBConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathDB + ";Jet OLEDB:Database Password=humus;";
                flag = true;
            }
            else if (File.Exists("nbu.ini"))
            {
                try
                {
                    pathDB = File.ReadAllText("nbu.ini", Encoding.GetEncoding(1251));
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Файл настроек недоступен для чтения " + ex.Message);
                }
                if (File.Exists(pathDB))
                {
                    strDBConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + @pathDB + ";Jet OLEDB:Database Password=humus;";
                    flag = true;
                }
            }
        }
//=================================================================================================================
// Загрузка формы, заполнение первого грида, формирование таблицы второго грида
        private void Form1_Load(object sender, EventArgs e)
        {
            label6.Text = DateTime.Today.ToString("D");
            if (flag)
            {
                ConnectDB(strDBConnect);
                //---------------
                //button1.Enabled = false;
                //---------------
                dataGridView1.Sort(dataGridView1.Columns[2], ListSortDirection.Ascending);
                datagridview2();
                dataGridView2.AllowUserToAddRows = false;
                dataGridView2.Refresh();
                dataGridView2.Columns[5].Visible = true;
                dataGridView2.Columns[6].Visible = true;
                dataGridView2.Columns[16].Visible = true;
                dataGridView2.Columns[17].Visible = true;
                dataGridView2.Columns[5].HeaderText = "Дата рождения";
                dataGridView2.Columns[6].HeaderText = "Отдел";
                dataGridView2.Columns[16].HeaderText = "Будет, лет";
                dataGridView2.Columns[17].HeaderText = "Осталось, дней";
                dataGridView2.Columns[2].Width = 130;
                dataGridView2.Columns[3].Width = 130;
                dataGridView2.Columns[4].Width = 130;
                dataGridView2.Columns[5].Width = 110;
                dataGridView2.Columns[6].Width = 180;
                dataGridView2.Columns[16].Width = 48;
                dataGridView2.Columns[17].Width = 59;
                dataGridView2.Columns[1].Visible = false;
                dataGridView2.Columns[8].Visible = false;
                dataGridView2.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView2.Columns[16].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView2.Columns[17].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView2.Sort(dataGridView2.Columns[17], ListSortDirection.Ascending);
                dataGridView2.Columns[5].DefaultCellStyle.Format = "D";
                dataGridView2.Columns[6].DisplayIndex = 5;
                dataGridView2.DefaultCellStyle.SelectionBackColor = Color.Peru;
                for (int i = 0; i < dataGridView2.Columns.Count; i++)
                    dataGridView2.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            if (dataGridView1.RowCount > 0)
            {
                вырезатьToolStripMenuItem.Enabled = true;
                редактироватьToolStripMenuItem.Enabled = true;
                пингToolStripMenuItem.Enabled = true;
                полнаяToolStripMenuItem1.Enabled = true;
                сохранитьToolStripMenuItem.Enabled = true;
            }
            
        }

//=================================================================================================================
// Подключение к базе
        private void ConnectDB(string strDBConnect)
        {
            try
            {
                cn = new OleDbConnection(strDBConnect);
                cn.Open();
                ad = new OleDbDataAdapter(strSQL, cn);
                dSet = new DataSet();
                ad.Fill(dSet);
                dSet.Tables[0].Columns.Add("years", typeof(int));
                dSet.Tables[0].Columns.Add("days", typeof(int));
                dataGridView1.DataSource = dSet.Tables[0].DefaultView;
                dataGridView1.Columns[0].Visible = false;
                dataGridView1.Columns[1].HeaderText = "Таб. №";
                dataGridView1.Columns[2].HeaderText = "Фамилия";
                dataGridView1.Columns[3].HeaderText = "Имя";
                dataGridView1.Columns[4].HeaderText = "Отчество";
                dataGridView1.Columns[8].HeaderText = "Телефон";
                dataGridView1.Columns[1].Width = 66;
                dataGridView1.Columns[2].Width = 160;
                dataGridView1.Columns[3].Width = 150;
                dataGridView1.Columns[4].Width = 150;
                dataGridView1.Columns[8].Width = 50;
                dataGridView1.Columns[5].Visible = false;
                dataGridView1.Columns[6].Visible = false;
                dataGridView1.Columns[7].Visible = false;
                dataGridView1.Columns[9].Visible = false;
                dataGridView1.Columns[10].Visible = false;
                dataGridView1.Columns[11].Visible = false;
                dataGridView1.Columns[12].Visible = false;
                dataGridView1.Columns[13].Visible = false;
                dataGridView1.Columns[14].Visible = false;
                dataGridView1.Columns[15].Visible = false;
                dataGridView1.Columns[16].Visible = false;
                dataGridView1.Columns[17].Visible = false;
                dataGridView1.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                cn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Данные не могут быть прочитаны из базы. " + ex.Message);
            }
        }
//=================================================================================================================
// Заполнение данными второго грида
        private void datagridview2()
        {
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                DateTime birthday = (DateTime)dataGridView1[5, i].Value;
                DateTime today = DateTime.Today;
                DateTime futureBirthday = birthday.AddYears(today.Year - birthday.Year);
                if (futureBirthday < today)
                    futureBirthday = birthday.AddYears(today.Year - birthday.Year + 1);
                dataGridView1[17, i].Value = (futureBirthday - today).TotalDays;
                DateTime qqq = (DateTime)dataGridView1[5, i].Value;
                if (dataGridView1[17, i].Value.ToString() == "0")
                    dataGridView1[16, i].Value = (CalculateAge(qqq)).ToString();
                else
                    dataGridView1[16, i].Value = (CalculateAge(qqq) + 1).ToString();
            }
            if (dataGridView2.Columns.Count == 0)
            {
                foreach (DataGridViewColumn dgvc in dataGridView1.Columns)
                {
                    dataGridView2.Columns.Add(dgvc.Clone() as DataGridViewColumn);
                }
            }
            DataGridViewRow row = new DataGridViewRow();
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (Convert.ToInt32(dataGridView1.Rows[i].Cells[17].Value) <= numericUpDown1.Value)
                {
                    row = (DataGridViewRow)dataGridView1.Rows[i].Clone();
                    int intColIndex = 0;
                    foreach (DataGridViewCell cell in dataGridView1.Rows[i].Cells)
                    {
                        row.Cells[intColIndex].Value = cell.Value;
                        intColIndex++;
                    }
                    dataGridView2.Rows.Add(row);
                }
            }

            for (int i = 0; i < dataGridView2.RowCount; i++)
            {
                if ((int)dataGridView2[16, i].Value % 5 == 0)
                {
                    dataGridView2.Rows[i].Cells[2].Style.ForeColor = Color.Red;
                    dataGridView2.Rows[i].Cells[3].Style.ForeColor = Color.Red;
                    dataGridView2.Rows[i].Cells[4].Style.ForeColor = Color.Red;
                    dataGridView2.Rows[i].Cells[5].Style.ForeColor = Color.Red;
                    dataGridView2.Rows[i].Cells[6].Style.ForeColor = Color.Red;
                    dataGridView2.Rows[i].Cells[16].Style.ForeColor = Color.Red;
                    dataGridView2.Rows[i].Cells[17].Style.ForeColor = Color.Red;
                }
            }
            //dataGridView2.Sort(dataGridView2.Columns[17], ListSortDirection.Ascending);
            if (dataGridView2.RowCount > 0)
            {
                dataGridView2.CurrentCell = dataGridView2[2, 0];
            }
        }
//=================================================================================================================
//Предупреждение, если файл БД не найден 
        private void Form1_Shown(object sender, System.EventArgs e) 
        {
            if (!flag)
                MessageBox.Show("Файл базы данных не найден!\nНайдите его через  Меню -> Путь к базе данных ", "Внимание!!!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
//=================================================================================================================
//Изменение обоих Датагридов после Добавления, Изменения или Удаления данных, скрытия ненужных пунктов меню для записи
        private void Form1_Activated(object sender, System.EventArgs e)
        {
                int rowIndex = -1;
                switch (GlobalData.prizn)
                {
                    case 1:
                        string query2 = "Select Max(id) FROM Table1";
                        cmd = new OleDbCommand(GlobalData.querySQL, cn);
                        if (GlobalData.pictureBox != null)
                        {
                            MemoryStream Mystream = new MemoryStream();
                            Bitmap BmpImage = new Bitmap(GlobalData.pictureBox);
                            BmpImage.Save(Mystream, System.Drawing.Imaging.ImageFormat.Jpeg);
                            byte[] ImageAsBytes = Mystream.ToArray();
                            imageParametr = cmd.Parameters.AddWithValue("@Foto", SqlDbType.Binary);
                            imageParametr.Value = ImageAsBytes;
                            imageParametr.Size = ImageAsBytes.Length;
                            Mystream.Dispose();
                        }
                        else
                            imageParametr = cmd.Parameters.AddWithValue("@Foto", DBNull.Value);
                        try
                        {
                            cn.Open();
                            cmd.ExecuteNonQuery();
                            cmd.CommandText = query2;
                            maxid = (Int32)cmd.ExecuteScalar();
                            cn.Close();
                        }
                        catch (Exception ex)
                        {
                            //cn.Close();
                            MessageBox.Show("Запись в базу не выполнена! " + ex.Message);
                        }
                        dSet.Clear();
                        ad.Fill(dSet);
                        dataGridView2.Rows.Clear();
                        datagridview2();
                        foreach (DataGridViewRow row in dataGridView1.Rows)
                        {
                            if (row.Cells[0].Value.Equals(maxid))
                            {
                                rowIndex = row.Index;
                                break;
                            }
                        }
                        dataGridView1.CurrentCell = dataGridView1[1, rowIndex];
                        break;
                    case 2:
                        cmd = new OleDbCommand("delete from Table1 where id=" + GlobalData.id, cn);
                        try
                        {
                            cn.Open();
                            cmd.ExecuteNonQuery();
                            cn.Close();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Удаление из базы не выполнено! " + ex.Message);
                        }
                        dSet.Clear();
                        ad.Fill(dSet);
                        dataGridView2.Rows.Clear();
                        datagridview2();
                        break;
                    case 3:
                        cmd = new OleDbCommand(GlobalData.querySQL, cn);
                        if (GlobalData.pictureBox != null)
                        {
                            MemoryStream Mystream = new MemoryStream();
                            Bitmap BmpImage = new Bitmap(GlobalData.pictureBox);
                            BmpImage.Save(Mystream, System.Drawing.Imaging.ImageFormat.Jpeg);
                            byte[] ImageAsBytes = Mystream.ToArray();
                            imageParametr = cmd.Parameters.AddWithValue("@Foto", SqlDbType.Binary);
                            imageParametr.Value = ImageAsBytes;
                            imageParametr.Size = ImageAsBytes.Length;
                            Mystream.Dispose();
                        }
                        else
                            imageParametr = cmd.Parameters.AddWithValue("@Foto", DBNull.Value);
                        try
                        {
                            cn.Open();
                            cmd.ExecuteNonQuery();
                            cn.Close();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Запись в базу не выполнена! " + ex.Message);
                        }
                        dSet.Clear();
                        ad.Fill(dSet);
                        dataGridView2.Rows.Clear();
                        datagridview2();
                        foreach (DataGridViewRow row in dataGridView1.Rows)
                        {
                            if (row.Cells[0].Value.Equals(GlobalData.id))
                            {
                                rowIndex = row.Index;
                                break;
                            }
                        }
                        dataGridView1.CurrentCell = dataGridView1[1, rowIndex];
                        break;
                    case 4:
                        foreach (DataGridViewRow row in dataGridView1.Rows)
                        {
                            if (row.Cells[0].Value.Equals(GlobalData.id))
                            {
                                rowIndex = row.Index;
                                break;
                            }
                        }
                        dataGridView1.CurrentCell = dataGridView1[1, rowIndex];
                        break;
                    //default:
                    //    if (dataGridView1.RowCount > 0)
                    //        dataGridView1.CurrentCell = dataGridView1[1, 0];
                    //    break;
                }
                GlobalData.prizn = 0;
                if (dataGridView1.RowCount == 0)
                {
                    вырезатьToolStripMenuItem.Enabled = false;
                    редактироватьToolStripMenuItem.Enabled = false;
                    пингToolStripMenuItem.Enabled = false;
                    полнаяToolStripMenuItem1.Enabled = false;
                    сохранитьToolStripMenuItem.Enabled = false;
                    pictureBox1.Image = null;
                }
                else
                {
                    вырезатьToolStripMenuItem.Enabled = true;
                    редактироватьToolStripMenuItem.Enabled = true;
                    пингToolStripMenuItem.Enabled = true;
                    полнаяToolStripMenuItem1.Enabled = true;
                    сохранитьToolStripMenuItem.Enabled = true;
                }
                
            }
          
//=================================================================================================================
// Включение проигрования мелодии
        private void sound(string myAudio1)
        {
            try
            {
                File.WriteAllBytes(myAudio, Properties.Resources.sound);
                Media.Player.GetPlayer().Play(myAudio);
                this.btnPlay.Image = global::PeopleNBU.Properties.Resources.Pause;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Проблемы с аудио.\nПроверьте наличие аудио кодеков " + ex.Message);
            }
        }

//=================================================================================================================
// Вычисление возраста сотрудника        
        public static int CalculateAge(DateTime BirthDate)
        {
            int YearsPassed = DateTime.Now.Year - BirthDate.Year;
            if (DateTime.Now.Month < BirthDate.Month || (DateTime.Now.Month == BirthDate.Month && DateTime.Now.Day < BirthDate.Day))
            {
                YearsPassed--;
            }
            return YearsPassed;
        }
//=================================================================================================================     
// Play/Pause плеера
        private void btnPlay_Click(object sender, EventArgs e)
        {
            if (Media.Player.GetPlayer().IsPaused())
            {
                Media.Player.GetPlayer().Play(true);
                this.btnPlay.Image = global::PeopleNBU.Properties.Resources.Pause;
            }
            else if (Media.Player.GetPlayer().IsPlaying())
            {
                Media.Player.GetPlayer().Pause();
                this.btnPlay.Image = global::PeopleNBU.Properties.Resources.Play;
            }
        }
//=================================================================================================================
// Управление громкостью плеера
        private void trackBarValume_Scroll(object sender, EventArgs e) 
        {
            if (Media.Player.GetPlayer().IsPlaying())
            {
                Media.Player.GetPlayer().MasterVolume = trackBarValume.Value;
            }
        }
//=================================================================================================================
// Закрытие главной формы 
        private void Form1_FormClosing(Object sender, FormClosingEventArgs e)
        {
            if (Media.Player.GetPlayer().IsOpen())
            {
                Media.Player.GetPlayer().Close();
            }
            if (File.Exists(myAudio))
                File.Delete(myAudio);
            if (File.Exists(Path.GetTempPath() + "Help.chm"))
                File.Delete(Path.GetTempPath() + "Help.chm");
            Application.Exit();
        }
//=================================================================================================================
//Меню Выход
        private void выходToolStripMenuItem_Click(object sender, EventArgs e) 
        {
            Application.Exit();
            Dispose();
        }
//=================================================================================================================
// Поиск файла БД
        private void открытьToolStripMenuItem_Click(object sender, EventArgs e) 
        {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                openFileDialog1.InitialDirectory = "D:\\";
                openFileDialog1.Filter = "DB files (*.accdb)|*.accdb";
                openFileDialog1.Title = "Поиск файла базы данных";
                openFileDialog1.RestoreDirectory = true;
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        pathDB = openFileDialog1.FileName;
                        File.WriteAllText("nbu.ini", pathDB, Encoding.GetEncoding(1251));
                        strDBConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + @pathDB + ";Persist Security Info=False";
                        flag = true;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Файл БД не может быть прочитан: " + ex.Message);
                    }
                }
     Form1_Load(sender,e);
         }
//=================================================================================================================
//Поиск по табельному номеру
         private void textBox1_TextChanged(object sender, EventArgs e) 
        {
            dataGridView1.Sort(dataGridView1.Columns[1], ListSortDirection.Ascending);
            for (int i = 0; i < dataGridView1.RowCount; i++)
                if (dataGridView1[1, i].Value.ToString() != "")
                {
                    if (dataGridView1[1, i].FormattedValue.ToString().Substring(0, textBox1.Text.Length).Equals(textBox1.Text.Trim(), StringComparison.CurrentCultureIgnoreCase))
                    {
                        dataGridView1.Rows[i].Cells[1].Selected = true;
                        return;
                    }
                }
         }
//=================================================================================================================
// Выход из поиска по табелю
         private void textBox1_LostFocus(object sender, System.EventArgs e) 
         {
             if (p3)
             {
                 this.textBox1.TextChanged -= new System.EventHandler(this.textBox1_TextChanged);
                 textBox1.Clear();
                 //dataGridView1.Sort(dataGridView1.Columns[2], ListSortDirection.Ascending);
                this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
             }
             p3 = true;
         }
//=================================================================================================================
// Изменение рассскладки в поле поиска по фамилии
         private void textBox2_GotFocus(object sender, System.EventArgs e) 
         {
             InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(new System.Globalization.CultureInfo("uk-UA")); 
         }
//=================================================================================================================
// Поиск по фамилии
         private void textBox2_TextChanged(object sender, EventArgs e) 
         {
             dataGridView1.Sort(dataGridView1.Columns[2], ListSortDirection.Ascending);
             try
             {
                 for (int i = 0; i < dataGridView1.RowCount; i++)
                 {
                     if (dataGridView1[2, i].FormattedValue.ToString().Length >= textBox2.Text.Trim().Length) 
                     {
                         if (dataGridView1[2, i].FormattedValue.ToString().Substring(0, textBox2.Text.Length).Equals(textBox2.Text.Trim(), StringComparison.CurrentCultureIgnoreCase))
                         {
                             dataGridView1.Rows[i].Cells[2].Selected = true;
                             return;
                         }
                     }
                 }
             }
             catch (Exception ex)
             {
                 MessageBox.Show("Ошибка: " + ex.Message);
             }
         }
//=================================================================================================================
// Выход из поиска по фамилии
         private void textBox2_LostFocus(object sender, System.EventArgs e) 
         {
             if (p3)
             {
                 this.textBox2.TextChanged -= new System.EventHandler(this.textBox2_TextChanged);
                 textBox2.Clear();
                 this.textBox2.TextChanged += new System.EventHandler(this.textBox2_TextChanged);
             }
             p3 = true;
         }
//=================================================================================================================
// Условия ввода в поле поиска по табелю
         private void textBox1_KeyPress(object sender, KeyPressEventArgs e) 
         {
             const char Delete = (char)8;
             if (e.KeyChar.Equals((char)0x1b))
             {
                 e.Handled = true;
                 textBox1.Clear();
             }
             else if (e.KeyChar.Equals((char)13))
             {
                 e.Handled = true;
                 this.SelectNextControl((Control)sender, true, true, true, true);
             }
             else if (!Char.IsDigit(e.KeyChar) && e.KeyChar != Delete)
             {
                 e.Handled = true;
                 p3 = false;
                 MessageBox.Show("В это поле для поиска необходимо вводить только цифры!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
             }
         }
//=================================================================================================================
// Условия ввода в поле поиска по фамилии
         private void textBox2_KeyPress(object sender, KeyPressEventArgs e) 
         {
             const char Delete = (char)8;
             if (e.KeyChar.Equals((char)0x1b))
             {
                 e.Handled = true;
                 textBox2.Clear();
             }
             else if (e.KeyChar.Equals((char)13))
             {
                 e.Handled = true;
                 this.SelectNextControl((Control)sender, true, true, true, true);
             }
             else if (Regex.Match(e.KeyChar.ToString(), @"[0-9]").Success && e.KeyChar != Delete)
             {
                 e.Handled = true;
                 p3 = false;
                 MessageBox.Show("В это поле для поиска необходимо вводить только буквы!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
             }
         }
//=================================================================================================================
// меню добавление нового сотрудника
         private void добавитьToolStripMenuItem_Click(object sender, EventArgs e) 
         {
             GlobalData.arrayTabNumber = new string[dataGridView1.RowCount];
             PerenosDannih();
             for (int i = 0; i < GlobalData.arrayTabNumber.Length; i++)
             {
                 GlobalData.arrayTabNumber[i] = dataGridView1[1, i].Value.ToString();
             }
             Form2 frm_Add = new Form2();
             frm_Add.ShowDialog();
             ObnulDannih();
         }
//=================================================================================================================
//Вывод фотографии в пикча-бокс
         private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e) 
         {
             
             if (dataGridView1[11, dataGridView1.CurrentRow.Index].Value == DBNull.Value)
             {
                 pictureBox1.Image = null;
                 pictureBox1.Invalidate();
             }
             else
             {
                 byte[] FetchedImgBytes = (byte[])dataGridView1[11, dataGridView1.CurrentRow.Index].Value;
                 MemoryStream stream = new MemoryStream(FetchedImgBytes);
                 pictureBox1.Image = Image.FromStream(stream);
                
             }
             if (dataGridView1[9, dataGridView1.CurrentRow.Index].Value.ToString().Trim().Length == 0)
                 пингToolStripMenuItem.Enabled = false;
             else
             {
                 пингToolStripMenuItem.Enabled = true;
                 ip = dataGridView1[9, dataGridView1.CurrentRow.Index].Value.ToString().Trim();
             }
         }
//=================================================================================================================
// Если фотография сотрудника отсутствует
         private void pictureBox1_Paint(object sender, PaintEventArgs e)
         {
             if (pictureBox1.Image == null)
             {
                 using (Font myFont = new Font("Arial", 9))
                 {
                     e.Graphics.DrawString("Фотография", myFont, Brushes.Green, new Point(47, 70));
                     e.Graphics.DrawString("отсутствует", myFont, Brushes.Green, new Point(50, 90));
                 }
             }
         }
//=================================================================================================================
// Перенос значений полей записи между формами
         private void PerenosDannih()
         {
             GlobalData.id = (Int32)dataGridView1[0, dataGridView1.CurrentRow.Index].Value;
             GlobalData.Tab = dataGridView1[1, dataGridView1.CurrentRow.Index].Value.ToString();
             GlobalData.F = dataGridView1[2, dataGridView1.CurrentRow.Index].Value.ToString();
             GlobalData.I = dataGridView1[3, dataGridView1.CurrentRow.Index].Value.ToString();
             GlobalData.O = dataGridView1[4, dataGridView1.CurrentRow.Index].Value.ToString();
             GlobalData.Bithday = (DateTime)dataGridView1[5, dataGridView1.CurrentRow.Index].Value;
             GlobalData.Otdel = dataGridView1[6, dataGridView1.CurrentRow.Index].Value.ToString();
             GlobalData.Dolzhn = dataGridView1[7, dataGridView1.CurrentRow.Index].Value.ToString();
             GlobalData.Tel = dataGridView1[8, dataGridView1.CurrentRow.Index].Value.ToString();
             GlobalData.Comp = dataGridView1[9, dataGridView1.CurrentRow.Index].Value.ToString();
             GlobalData.Email = dataGridView1[10, dataGridView1.CurrentRow.Index].Value.ToString();
             GlobalData.pictureBox = pictureBox1.Image;
             GlobalData.Adress = dataGridView1[12, dataGridView1.CurrentRow.Index].Value.ToString();
             GlobalData.BIRTHPLACE = dataGridView1[13, dataGridView1.CurrentRow.Index].Value.ToString();
             GlobalData.Passport = dataGridView1[14, dataGridView1.CurrentRow.Index].Value.ToString();
             GlobalData.IPN = dataGridView1[15, dataGridView1.CurrentRow.Index].Value.ToString();
         }
//=================================================================================================================
// Обнуление глобальных переменных
         private void ObnulDannih()
         {
             GlobalData.id = 0;
             GlobalData.Tab = " ";
             GlobalData.F = " ";
             GlobalData.I = " ";
             GlobalData.O = " ";
             GlobalData.Bithday = DateTime.Today; 
             GlobalData.Otdel = " ";
             GlobalData.Dolzhn = " ";
             GlobalData.Comp = " ";
             GlobalData.Tel = " ";
             GlobalData.Email = " ";
             GlobalData.pictureBox = null;
             GlobalData.Adress = " ";
             GlobalData.BIRTHPLACE = " ";
             GlobalData.Passport = " ";
             GlobalData.IPN = " ";
         }
//=================================================================================================================  
// меню удаления сотрудника
        private void вырезатьToolStripMenuItem_Click(object sender, EventArgs e) 
         {
             PerenosDannih();
             Form3 frm_Del = new Form3();
             frm_Del.ShowDialog();
             ObnulDannih();
         }
//=================================================================================================================
// меню редактирования данных сотрудника
         private void редактироватьToolStripMenuItem_Click(object sender, EventArgs e)
         {
             PerenosDannih();
             GlobalData.arrayTabNumber = new string[dataGridView1.RowCount];
             for (int i = 0; i < GlobalData.arrayTabNumber.Length; i++)
             {
                 GlobalData.arrayTabNumber[i] = dataGridView1[1, i].Value.ToString();
             }
             Form4 frm_Red = new Form4();
             frm_Red.ShowDialog();
             ObnulDannih();
         }
//==============================================================================================
// Синхронизация с Юбиллеем
         private void синхронизацияБДToolStripMenuItem_Click(object sender, EventArgs e)
         {
             if (!File.Exists(@"k:\BD_SHEF\table0.DB"))
                MessageBox.Show("Синхронизация НЕ выполнена\n\nПроверьте сетевое подключение!\n\nНет доступа к диску (К)", "Внимание!", MessageBoxButtons.OK,
                MessageBoxIcon.Error);
             else
             {
                try
                {
                    File.Copy(@"k:\BD_SHEF\table0.DB", Path.GetTempPath() + @"\table0.DB", true);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString(),  "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                OleDbConnection cn2;
                //string strDBConnect2 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d://My Documents//Табельные номера//;Extended Properties=Paradox 5.x";
                //string strDBConnect2 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d://tmp//Tabel//;Extended Properties=dBASE IV";

                string strDBConnect2 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Path.GetTempPath() + ";Extended Properties=Paradox 5.x";

                cn2 = new OleDbConnection(strDBConnect2);
                OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT * FROM table0", cn2);
                DataTable table = new DataTable();
                try
                {
                    adapter.Fill(table);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка подключения к файлу " + Path.GetTempPath() + @"\table0.DB " + ex.ToString(), "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                string[,] Array = new string[table.Rows.Count, 6];
                 
//ФИО в одном элементе массива и должность с отделом тоже в одном
                int i = 0;
                foreach (DataRow row in table.Rows)
                {
                    Array[i, 0] = row[0].ToString();
                    Array[i, 3] = row[1].ToString();
                    Array[i, 5] = row[4].ToString();
                    i++;
                }
                //Разносим ФИО по разным элементам массива
                string[] words;
                try
                {
                for (i = 0; i < Array.GetLength(0); i++)
                {
                    words = Array[i, 0].Split(' ');
                    Array[i, 0] = words[0].Replace("'", "''");
                    Array[i, 0] = Array[i, 0].Replace("I", "І");
                    Array[i, 0] = Array[i, 0].Replace("ґ", "є");
                    Array[i, 0] = Array[i, 0].Replace("i", "і");
                    Array[i, 0] = Array[i, 0].Replace("Ґ", "Є");
                    Array[i, 0] = Array[i, 0].Replace("\"", "'");
                    Array[i, 0] = Array[i, 0].Replace("''", "'");
                    Array[i, 1] = words[1].Replace("'", "''");
                    Array[i, 1] = Array[i, 1].Replace("I", "І");
                    Array[i, 1] = Array[i, 1].Replace("є", "ї");
                    Array[i, 1] = Array[i, 1].Replace("Є", "Ї");
                    Array[i, 1] = Array[i, 1].Replace("ґ", "є");
                    Array[i, 1] = Array[i, 1].Replace("i", "і");
                    Array[i, 1] = Array[i, 1].Replace("Ґ", "Є");
                    Array[i, 1] = Array[i, 1].Replace("\"", "'");
                    Array[i, 1] = Array[i, 1].Replace("''", "'");
                    Array[i, 2] = words[2].Replace("'", "''");
                    Array[i, 2] = Array[i, 2].Replace("I", "І");
                    Array[i, 2] = Array[i, 2].Replace("є", "ї");
                    Array[i, 2] = Array[i, 2].Replace("Є", "Ї");
                    Array[i, 2] = Array[i, 2].Replace("ґ", "є");
                    Array[i, 2] = Array[i, 2].Replace("i", "і");
                    Array[i, 2] = Array[i, 2].Replace("Ґ", "Є");
                    Array[i, 2] = Array[i, 2].Replace("\"", "'");
                    Array[i, 2] = Array[i, 2].Replace("''", "'");
                }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
//Разносим должность и отдел по разным элементам массива
                for (i = 0; i < Array.GetLength(0); i++)
                {
                    int index = 0;
                    foreach (char c in Array[i, 3])  
                    {
                        if (char.IsLower(c))
                        {
                            index = Array[i, 3].IndexOf(c);

                            if (index < 5)
                            {
                                string x = Array[i, 3];
                                index = (from ch in x.ToArray()
                                        where Char.IsUpper(ch)
                                        select x.LastIndexOf(ch)).Last() + 1;
                            }
                            break;
                        }
                    }
                    Array[i, 4] = Array[i, 3].Substring(index - 1).Replace("'", "''");
                    Array[i, 4] = Array[i, 4].Replace("є", "ї");
                    Array[i, 4] = Array[i, 4].Replace("I", "І");
                    Array[i, 4] = Array[i, 4].Replace("i", "і");

                    string s = Array[i, 3].Remove(index - 2).Replace("'", "''");
                    Array[i, 3] = s.Substring(0, 1) + s.Substring(1).ToLower();
                    Array[i, 3] = Array[i, 3].Replace("є", "ї");
                    Array[i, 3] = Array[i, 3].Replace("I", "І");
                    Array[i, 3] = Array[i, 3].Replace("i", "і"); 
                }

                cn = new OleDbConnection(strDBConnect);
                cn.Open();
                int k;
                bool k1 = true;
                string new_p = "\n\nНовые сотрудники:\n\n";
                if (dataGridView1.RowCount == 0)
                    k = 0;
                else k = 1;
                for (int w = 0; w < Array.GetLength(0); w++)
                {
                    if (k == 1)
                    {
                        bool prizn = true;
                        for (int j = 0; j < dataGridView1.RowCount; j++)
                        {
                            if ((dataGridView1.Rows[j].Cells[2].Value.ToString() == Array[w, 0]) & (dataGridView1.Rows[j].Cells[5].Value.ToString() == Array[w, 5]))
                            {
                                prizn = false;
                                break;
                            }
                        }
                        if (prizn)
                        {
                            InsertSynhr(Array[w, 0], Array[w, 1], Array[w, 2], Array[w, 5], Array[w, 3], Array[w, 4]);
                            new_p = new_p + Array[w, 0] + " " + Array[w, 1] + " " + Array[w, 2] + "\n";
                            k1 = false;
                        }
                        
                    }
                    else
                    {
                        InsertSynhr(Array[w, 0], Array[w, 1], Array[w, 2], Array[w, 5], Array[w, 3], Array[w, 4]);
                        new_p = " ";
                        k1 = false;
                    }
                }
                if (k1)
                    new_p = "\n\nНовых сотрудников НЕТ.";

                cn.Close();
                Form1_Load(sender, e);
                MessageBox.Show("Синхронизация закончена " + new_p, "Внимание!");
                if (File.Exists(Path.GetTempPath() + @"\table0.DB"))
                {
                File.Delete(Path.GetTempPath() + @"\table0.DB");
                }
             }
         }
//==============================================================================================
// Непосредственный перенос из юбиллея в базу
         public void InsertSynhr(string fam, string ima, string otch, string date, string otd, string dol )
         {
             DateTime date1 = DateTime.Parse(date);
             string sql = "INSERT INTO Table1 (F,I,O,Bithday,Otdel,DOLZHN) VALUES(@1,@2,@3,@4,@5,@6)";
             using (OleDbCommand cmd = new OleDbCommand(sql, cn))
             {
                cmd.Parameters.AddWithValue("@1", fam);
                cmd.Parameters.AddWithValue("@2", ima);
                cmd.Parameters.AddWithValue("@3", otch);
                cmd.Parameters.AddWithValue("@4", date1);
                cmd.Parameters.AddWithValue("@5", otd);
                cmd.Parameters.AddWithValue("@6", dol);
                cmd.ExecuteNonQuery();
             }
         }
//==============================================================================================
// Меню полная инфа о сотруднике
         private void полнаяToolStripMenuItem1_Click(object sender, EventArgs e)
         {
            PerenosDannih();
            Form5 frm_All = new Form5();
            frm_All.ShowDialog();
            ObnulDannih();
         }
//==============================================================================================
// Меню пинг компа сотрудника, если он есть
         private void пингToolStripMenuItem_Click(object sender, EventArgs e)
         {
             Ping ping = new Ping();
             try
             {
                PingReply pingReply = ping.Send(ip);
                if (pingReply.Status == IPStatus.Success)
                {
                    MessageBox.Show("Человек на рабочем месте", "Комп включён", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                else
                {
                    MessageBox.Show("Компьютер выключен");
                }
             }
             catch 
             {
                 //MessageBox.Show(ex.InnerException.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                 MessageBox.Show("Компьютер выключен, либо ошибка определения имени в домене", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
             }
         }
//==============================================================================================
// Изменение за сколько дней смотреть дни рождения и перерсчет грида2
         private void numericUpDown1_TextChanged(object sender, EventArgs e)
         {
             dataGridView2.Rows.Clear();
             //foreach (DataGridViewColumn dgvc in dataGridView1.Columns)
             //    {
             //        dataGridView2.Columns.Add(dgvc.Clone() as DataGridViewColumn);
             //    }
            DataGridViewRow row = new DataGridViewRow();
             for (int i = 0; i < dataGridView1.Rows.Count; i++)
             {
                 //if (dataGridView1.Rows[i].Cells[1].Value.ToString() != "")
                 //MessageBox.Show(dataGridView1.Rows[i].Cells[17].Value.ToString());
                if (Convert.ToInt32(dataGridView1.Rows[i].Cells[17].Value) <= numericUpDown1.Value)
                {
                    row = (DataGridViewRow)dataGridView1.Rows[i].Clone();
                    int intColIndex = 0;
                    foreach (DataGridViewCell cell in dataGridView1.Rows[i].Cells)
                    {
                        row.Cells[intColIndex].Value = cell.Value;
                        intColIndex++;
                    }
                    dataGridView2.Rows.Add(row);
                }
             }
             if (dataGridView2.Columns.Count != 0)
                 dataGridView2.Sort(dataGridView2.Columns[17], ListSortDirection.Ascending);

             if ((numericUpDown1.Value == 2) || (numericUpDown1.Value == 3) || (numericUpDown1.Value == 4))
             {
                 label4.Text = "дня";
                 label3.Text = "На ближайшие";
             }
             else if (numericUpDown1.Value == 1)
             {
                 label4.Text = "день";
                 label3.Text = "На ближайший";
             }
             else
             {
                 label4.Text = "дней";
                 label3.Text = "На ближайшие";
             }
             for (int i = 0; i < dataGridView2.RowCount; i++)
             {
                 if ((int)dataGridView2[16, i].Value % 5 == 0)
                 {
                     dataGridView2.Rows[i].Cells[2].Style.ForeColor = Color.Red;
                     dataGridView2.Rows[i].Cells[3].Style.ForeColor = Color.Red;
                     dataGridView2.Rows[i].Cells[4].Style.ForeColor = Color.Red;
                     dataGridView2.Rows[i].Cells[5].Style.ForeColor = Color.Red;
                     dataGridView2.Rows[i].Cells[6].Style.ForeColor = Color.Red;
                     dataGridView2.Rows[i].Cells[16].Style.ForeColor = Color.Red;
                     dataGridView2.Rows[i].Cells[17].Style.ForeColor = Color.Red;
                 }
             }
            }


//==============================================================================================
// Окно о программе
        private void опрограммеToolStripMenuItem_Click(object sender, EventArgs e)
         {
             AboutBox1 box = new AboutBox1();
             box.ShowDialog();
         }
//==============================================================================================
// Меню выгрузить запись или всю базу в файл
         private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
         { 
             PerenosDannih();
             Form6 frm_Save = new Form6();
             frm_Save.Owner = this;
             frm_Save.ShowDialog();
             ObnulDannih();
         }
//==============================================================================================
//  Меню печати
         private void печатьToolStripMenuItem_Click(object sender, EventArgs e)
         {
             printDocument1.DefaultPageSettings.Landscape = true;
             if ((dataGridView1.Rows.Count % 31) > 0)
            {
                 allPage = dataGridView1.Rows.Count / 31 + 1;
             }
             else
             {
                 allPage = dataGridView1.Rows.Count;
             }
             using (var dlg = new CoolPrintPreview.CoolPrintPreviewDialog())
             {
                 //MessageBox.Show("1-й    " + ((ToolStrip)dlg.Controls[1]).Items[0].Name);
                 ((ToolStrip)(dlg.Controls[1])).Items.RemoveAt(1); 
                 dlg.Document = this.printDocument1;
                 dlg.ShowDialog(this);
             }
         }

        bool firstPagePrinted = true;
//==============================================================================================
//  Вывод на предпросмотр каждой страницы
        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            Rectangle rc = e.MarginBounds;
            rc.Height = e.MarginBounds.Height + 100;
            int height = 50;
            int x = 5;

            StringFormat str = new StringFormat();
            str.Alignment = StringAlignment.Center;
            str.LineAlignment = StringAlignment.Center;

            #region Заголовок
            string text1 = "Реестр сотрудников\nНационального банка Украины в Харьковской области";
            if (firstPagePrinted)
            {
                using (Font font1 = new Font("Arial", 9, FontStyle.Bold, GraphicsUnit.Point))
                {
                    Rectangle rect1 = new Rectangle(5, 10, 1130, 30);
                    e.Graphics.DrawString(text1, font1, Brushes.Black, rect1, str);
                }
                firstPagePrinted = false;
            }
            #endregion

            Font font2 = new Font("Arial", 7, FontStyle.Bold, GraphicsUnit.Point);
            Font font3 = new Font("Arial", 7, FontStyle.Regular, GraphicsUnit.Point);
         
            #region Draw Column 1
            e.Graphics.FillRectangle(Brushes.LightGray, new Rectangle(x, 50, 40, dataGridView1.Rows[0].Height));
            e.Graphics.DrawRectangle(Pens.Black, x, 50, 40, dataGridView1.Rows[0].Height);
            e.Graphics.DrawString(dataGridView1.Columns[1].HeaderText, font2, Brushes.Black, new RectangleF(x, 50, 40, dataGridView1.Rows[0].Height), str);
            #endregion
            x = x + 40;
            #region Draw column 2
            e.Graphics.FillRectangle(Brushes.LightGray, new Rectangle(x, 50, 95, dataGridView1.Rows[0].Height));
            e.Graphics.DrawRectangle(Pens.Black, x, 50, 95, dataGridView1.Rows[0].Height);
            e.Graphics.DrawString(dataGridView1.Columns[2].HeaderText, font2, Brushes.Black, new RectangleF(x, 50, 95, dataGridView1.Rows[0].Height), str);
            #endregion
            x = x + 95;
            #region Draw column 3
            e.Graphics.FillRectangle(Brushes.LightGray, new Rectangle(x, 50, 95, dataGridView1.Rows[0].Height));
            e.Graphics.DrawRectangle(Pens.Black, x, 50, 95, dataGridView1.Rows[0].Height);
            e.Graphics.DrawString(dataGridView1.Columns[3].HeaderText, font2, Brushes.Black, new RectangleF(x, 50, 95, dataGridView1.Rows[0].Height), str);
            #endregion
            x = x + 95;
            #region Draw column 4
            e.Graphics.FillRectangle(Brushes.LightGray, new Rectangle(x, 50, 95, dataGridView1.Rows[0].Height));
            e.Graphics.DrawRectangle(Pens.Black, x, 50, 95, dataGridView1.Rows[0].Height);
            e.Graphics.DrawString(dataGridView1.Columns[4].HeaderText, font2, Brushes.Black, new RectangleF(x, 50, 95, dataGridView1.Rows[0].Height), str);
            #endregion
            x = x + 95;
            #region Draw column 5
            e.Graphics.FillRectangle(Brushes.LightGray, new Rectangle(x, 50, 60, dataGridView1.Rows[0].Height));
            e.Graphics.DrawRectangle(Pens.Black, x, 50, 60, dataGridView1.Rows[0].Height);
            e.Graphics.DrawString("Дата\nрождения", font2, Brushes.Black, new RectangleF(x, 50, 60, dataGridView1.Rows[0].Height), str);
            #endregion
            x = x + 60;
            #region Draw column 6
            e.Graphics.FillRectangle(Brushes.LightGray, new Rectangle(x, 50, 250, dataGridView1.Rows[0].Height));
            e.Graphics.DrawRectangle(Pens.Black, x, 50, 250, dataGridView1.Rows[0].Height);
            e.Graphics.DrawString("Отдел", font2, Brushes.Black, new RectangleF(x, 50, 250, dataGridView1.Rows[0].Height), str);
            #endregion
            x = x + 250;
            #region Draw column 7
            e.Graphics.FillRectangle(Brushes.LightGray, new Rectangle(x, 50, 185, dataGridView1.Rows[0].Height));
            e.Graphics.DrawRectangle(Pens.Black, x, 50, 185, dataGridView1.Rows[0].Height);
            e.Graphics.DrawString("Должность", font2, Brushes.Black, new RectangleF(x, 50, 185, dataGridView1.Rows[0].Height), str);
            #endregion
            x = x + 185;
            #region Draw column 8
            e.Graphics.FillRectangle(Brushes.LightGray, new Rectangle(x, 50, 50, dataGridView1.Rows[0].Height));
            e.Graphics.DrawRectangle(Pens.Black, x, 50, 50, dataGridView1.Rows[0].Height);
            e.Graphics.DrawString("Телефон", font2, Brushes.Black, new RectangleF(x, 50, 50, dataGridView1.Rows[0].Height), str);
            #endregion
            x = x + 50;
            #region Draw column 9
            e.Graphics.FillRectangle(Brushes.LightGray, new Rectangle(x, 50, 75, dataGridView1.Rows[0].Height));
            e.Graphics.DrawRectangle(Pens.Black, x, 50, 75, dataGridView1.Rows[0].Height);
            e.Graphics.DrawString("Имя\nкомпьютера", font2, Brushes.Black, new RectangleF(x, 50, 75, dataGridView1.Rows[0].Height), str);
            #endregion
            x = x + 75;
            #region Draw column 10
            e.Graphics.FillRectangle(Brushes.LightGray, new Rectangle(x, 50, 185, dataGridView1.Rows[0].Height));
            e.Graphics.DrawRectangle(Pens.Black, x, 50, 185, dataGridView1.Rows[0].Height);
            e.Graphics.DrawString("Email", font2, Brushes.Black, new RectangleF(x, 50, 185, dataGridView1.Rows[0].Height), str);
            #endregion
            
            string text2 = "Страница " + _page++ + " из " + allPage;
            using (Font font1 = new Font("Arial", 9, FontStyle.Regular, GraphicsUnit.Point))
            {
                Rectangle rect2 = new Rectangle(1000, 780, 135, 15);
                e.Graphics.DrawString(text2, font1, Brushes.Black, rect2, str);
            }

            while (v < dataGridView1.Rows.Count)
            {

                if (height > rc.Height)
                {
                    height = 50;
                    e.HasMorePages = true;
                    return;
                }

                str.Alignment = StringAlignment.Center;
                height += dataGridView1.Rows[v].Height;
                e.Graphics.DrawRectangle(Pens.Black, 5, height, 40, dataGridView1.Rows[0].Height);
                e.Graphics.DrawString(dataGridView1.Rows[v].Cells[1].FormattedValue.ToString(), font3, Brushes.Black, new RectangleF(5, height, 40, dataGridView1.Rows[0].Height), str);

                str.Alignment = StringAlignment.Near;
                e.Graphics.DrawRectangle(Pens.Black, 45, height, 95, dataGridView1.Rows[0].Height);
                e.Graphics.DrawString(dataGridView1.Rows[v].Cells[2].Value.ToString(), font3, Brushes.Black, new RectangleF(45, height, 95, dataGridView1.Rows[0].Height), str);

                e.Graphics.DrawRectangle(Pens.Black, 140, height, 95, dataGridView1.Rows[0].Height);
                e.Graphics.DrawString(dataGridView1.Rows[v].Cells[3].Value.ToString(), font3, Brushes.Black, new RectangleF(140, height, 95, dataGridView1.Rows[0].Height), str);

                e.Graphics.DrawRectangle(Pens.Black, 235, height, 95, dataGridView1.Rows[0].Height);
                e.Graphics.DrawString(dataGridView1.Rows[v].Cells[4].Value.ToString(), font3, Brushes.Black, new RectangleF(235, height, 95, dataGridView1.Rows[0].Height), str);

                str.Alignment = StringAlignment.Center;
                DateTime date1 = (DateTime)dataGridView1.Rows[v].Cells[5].Value;
                e.Graphics.DrawRectangle(Pens.Black, 330, height, 60, dataGridView1.Rows[0].Height);
                e.Graphics.DrawString(date1.ToShortDateString(), font3, Brushes.Black, new RectangleF(330, height, 60, dataGridView1.Rows[0].Height), str);

                str.Alignment = StringAlignment.Near;
                e.Graphics.DrawRectangle(Pens.Black, 390, height, 250, dataGridView1.Rows[0].Height);
                e.Graphics.DrawString(dataGridView1.Rows[v].Cells[6].Value.ToString(), font3, Brushes.Black, new RectangleF(390, height, 250, dataGridView1.Rows[0].Height), str);

                e.Graphics.DrawRectangle(Pens.Black, 640, height, 185, dataGridView1.Rows[0].Height);
                e.Graphics.DrawString(dataGridView1.Rows[v].Cells[7].Value.ToString(), font3, Brushes.Black, new RectangleF(640, height, 185, dataGridView1.Rows[0].Height), str);

                str.Alignment = StringAlignment.Center;
                e.Graphics.DrawRectangle(Pens.Black, 825, height, 50, dataGridView1.Rows[0].Height);
                e.Graphics.DrawString(dataGridView1.Rows[v].Cells[8].Value.ToString(), font3, Brushes.Black, new RectangleF(825, height, 50, dataGridView1.Rows[0].Height), str);

                e.Graphics.DrawRectangle(Pens.Black, 875, height, 75, dataGridView1.Rows[0].Height);
                e.Graphics.DrawString(dataGridView1.Rows[v].Cells[9].Value.ToString(), font3, Brushes.Black, new RectangleF(875, height, 75, dataGridView1.Rows[0].Height), str);

                str.Alignment = StringAlignment.Near;
                e.Graphics.DrawRectangle(Pens.Black, 950, height, 185, dataGridView1.Rows[0].Height);
                e.Graphics.DrawString(dataGridView1.Rows[v].Cells[10].Value.ToString(), font3, Brushes.Black, new RectangleF(950, height, 185, dataGridView1.Rows[0].Height), str);
                v++;
            }
            
        }
//==============================================================================================
//  Настройки вывода на печать каждой страницы
        void printDocument1_BeginPrint(object sender, System.Drawing.Printing.PrintEventArgs e)
        {
            v = 0;
            _page = 1;
            firstPagePrinted = true;
        }

        //==============================================================================================
        //  Меню печати
        private void предварительныйпросмотрToolStripMenuItem_Click(object sender, EventArgs e)
        {
            printDocument1.DefaultPageSettings.Landscape = true;
            if ((dataGridView1.Rows.Count % 31) > 0)
            {
                allPage = dataGridView1.Rows.Count / 31 + 1;
            }
            else
            {
                allPage = dataGridView1.Rows.Count;
            }
            using (var dlg = new CoolPrintPreview.CoolPrintPreviewDialog())
            {
                ((ToolStrip)(dlg.Controls[1])).Items.RemoveAt(1);

                ((ToolStrip)(dlg.Controls[1])).Items.RemoveAt(0);
                dlg.Document = this.printDocument1;
                dlg.ShowDialog(this);
            }
        }

        private void содержаниеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Process.Start(global::PeopleNBU.Properties.Resources.Help);
            File.WriteAllBytes(Path.GetTempPath() + "Help.chm", Properties.Resources.Help);
            Help.ShowHelp(this, Path.GetTempPath() + "Help.chm");
        }

        private void Form1_MouseUp(object sender, MouseEventArgs e)
        {
            foreach (Form form in Application.OpenForms)
            {
                if (form.Name == "Form8")
                {
                    form.Hide();
                }
            }
        }

        private void dataGridView1_MouseUp(object sender, MouseEventArgs e)
        {
            foreach (Form form in Application.OpenForms)
            {
                if (form.Name == "Form8")
                {
                    form.Hide();
                }
            }
        }

        private void pictureBox1_MouseUp(object sender, MouseEventArgs e)
        {
            foreach (Form form in Application.OpenForms)
            {
                if (form.Name == "Form8")
                {
                    form.Hide();
                }
            }
        }


        private void groupBox1_MouseUp(object sender, MouseEventArgs e)
        {
            foreach (Form form in Application.OpenForms)
            {
                if (form.Name == "Form8")
                {
                    form.Hide();
                }
            }
        } 
        
        
        private void dataGridView2_MouseUp(object sender, MouseEventArgs e)
        {
            DataGridView.HitTestInfo hti = dataGridView2.HitTest(e.X, e.Y);
            if (hti.Type == DataGridViewHitTestType.Cell)
            {
                dataGridView2.ClearSelection();
                dataGridView2.Rows[hti.RowIndex].Selected = true;
        
                if (dataGridView2[11, hti.RowIndex].Value == DBNull.Value)
                {
                    GlobalData1.pictureBox = null;
                }
                else
                {
                    byte[] FetchedImgBytes = (byte[])dataGridView2[11, hti.RowIndex].Value;
                    MemoryStream stream = new MemoryStream(FetchedImgBytes);
                    GlobalData1.pictureBox = Image.FromStream(stream);
                }
                Form8 foto = new Form8();

                if (e.Button == MouseButtons.Left)
                {
                    foreach (Form form in Application.OpenForms)
                    {
                        if (form.Name == "Form8")
                        {
                            form.Hide();
                        }
                    }
                }
                else if (e.Button == MouseButtons.Right)
                {
                    foreach (Form form in Application.OpenForms)
                    {
                        if (form.Name == "Form8")
                        {
                            form.Hide();
                        }
                    }
                    foto.Location = new Point(Cursor.Position.X, Cursor.Position.Y - 224);
                    foto.Show();
                }
            }
        }



    }
}