//==============================================================================================
// Форма редактирования информации о сотруднике
//==============================================================================================
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Data.OleDb;
using System.IO;

namespace PeopleNBU
{
    public partial class Form : Form
    {
        private string email, Comp;
        private TextBox txtBox1 = new TextBox();
        private TextBox txtBox2 = new TextBox();
        private TextBox txtBox3 = new TextBox();
        private TextBox txtBox4 = new TextBox();
        private TextBox txtBox5 = new TextBox();
        private TextBox txtBox6 = new TextBox();
        private TextBox txtBox7 = new TextBox();
        private TextBox txtBox8 = new TextBox();
        private TextBox txtBox9 = new TextBox();
        private int p1 = 0;
        private int p2 = 0;
       
        public Form()
        {
            InitializeComponent();
            textBox_Tab.Text = GlobalData.Tab;
            textBox_F.Text = GlobalData.F;
            textBox_I.Text = GlobalData.I;
            textBox_O.Text = GlobalData.O;
            dateTimePicker1.Value = GlobalData.Bithday;
            comboBox_Otdel.Text = GlobalData.Otdel;
            comboBox_Dolzhn.Text = GlobalData.Dolzhn;
            maskedTextBox_Tel.Text = GlobalData.Tel;
            textBox_Email.Text = GlobalData.Email;
            if (GlobalData.Email.Trim().Length != 0)
                textBox_Email.Text = GlobalData.Email.Remove(GlobalData.Email.Length - 17);
            else textBox_Email.Text = GlobalData.Email;
            pictureBox1.Image = GlobalData.pictureBox;
            
        }
//==============================================================================================
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
//==============================================================================================
// Кнопка "Отмена"
        private void button2_Click(object sender, EventArgs e)
        {
            if (p1 == 0)
            {
                GlobalData.prizn = 4;
                this.Close();
            }
            else if (p1 == 1)
            {
                button_Foto.Enabled = true;
                textBox_Tab.Visible = true;
                txtBox1.Visible = false;
                textBox_F.Visible = true;
                txtBox2.Visible = false;
                textBox_I.Visible = true;
                txtBox3.Visible = false;
                textBox_O.Visible = true;
                txtBox4.Visible = false;
                dateTimePicker1.Visible = true;
                txtBox5.Visible = false;
                comboBox_Otdel.Visible = true;
                txtBox6.Visible = false;
                comboBox_Dolzhn.Visible = true;
                txtBox7.Visible = false;
                txtBox8.Visible = false;
                txtBox9.Visible = false;
                label7.Visible = true;
                label9.Visible = true;
                label12.Visible = true;
                checkBox_Tel.Visible = true;
                maskedTextBox_Tel.Visible = true;
                checkBox_Comp.Visible = true;
                checkBox_Email.Visible = true;
                textBox_Email.Visible = true;
                label10.Visible = true;
                p1 = p2 = 0;
             }
          }
//==============================================================================================
// Условия ввода в поле табель
        private void textBox_Tab_KeyPress(object sender, KeyPressEventArgs e) 
        {
            const char Delete = (char)8;
            if (e.KeyChar.Equals((char)13))
            {
                e.Handled = true;
                this.SelectNextControl((Control)sender, true, true, true, true);
            }
            else if (!Char.IsDigit(e.KeyChar) && e.KeyChar != Delete)
            {
                e.Handled = true;
                MessageBox.Show("В это поле необходимо вводить только цифры!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
//==============================================================================================
// Проверка выхода из поля Табель
        private void textBox_Tab_Leave(object sender, EventArgs e)
        {
            if (button2.Focused)
                return;
            else
            {
                if (textBox_Tab.Text != GlobalData.Tab)
                {
                    if (textBox_Tab.Text.Length != 4)
                    {
                        MessageBox.Show("Номер табеля состоит из 4 цифр.",
                                "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
                        textBox_Tab.Focus();
                    }
                    else
                    {
                        for (int i = 0; i < GlobalData.arrayTabNumber.Length; i++)
                        {
                            if (GlobalData.arrayTabNumber[i] == textBox_Tab.Text)
                            {
                                textBox_Tab.Focus();
                                MessageBox.Show("Такой табельный номер существует! \n\n" + string.Format("{0,11}", " ") + "Будьте внимательны.",
                                "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
                                break;
                            }
                        }
                    }
                }
            }
        }
//==============================================================================================
//Инициализация перегрузок новых контролов, в которых редактируемая инфа о сотруднике 
        private void Form4_Load(object sender, EventArgs e)
        {
            if (maskedTextBox_Tel.Text.Trim() != "-")
            {
                checkBox_Tel.Checked = true;
                maskedTextBox_Tel.Enabled = true;
            }
            if (GlobalData.Comp.Trim().Length != 0)
            {
                checkBox_Comp.Checked = true;
            }
            if (textBox_Email.Text.Trim().Length == 0)
            {
                checkBox_Email.Checked = false;
                textBox_Email.Enabled = false;
            }
            else
            {
                checkBox_Email.Checked = true;
                checkBox_Email.Enabled = true;
                textBox_Email.Enabled = true;
            }
            if (pictureBox1.Image != null)
            {
                button3.Enabled = true;
            }

            TextBox[] arrText = new TextBox[] { textBox_F, textBox_I, textBox_O };
            foreach (TextBox i in arrText)
            {
                i.KeyPress += new KeyPressEventHandler(tb_KeyPress);
                i.GotFocus += new EventHandler(i_GotFocus);
                i.TextChanged += new EventHandler(i_TextChanged);
            }

            // Новые контролы
            this.Controls.Add(txtBox1);
            txtBox1.Location = new Point(11, 28);
            txtBox1.Size = new System.Drawing.Size(35, 20);
            
            this.Controls.Add(txtBox2);
            txtBox2.Location = new Point(11, 79);
            txtBox2.Size = new System.Drawing.Size(190, 21);
            
            this.Controls.Add(txtBox3);
            txtBox3.Location = new Point(11, 133);
            txtBox3.Size = new System.Drawing.Size(190, 21);
            
            this.Controls.Add(txtBox4);
            txtBox4.Location = new Point(11, 187);
            txtBox4.Size = new System.Drawing.Size(190, 21);
            
            this.Controls.Add(txtBox5);
            txtBox5.Location = new Point(11, 243);
            txtBox5.Size = new System.Drawing.Size(120, 21);
            
            this.Controls.Add(txtBox6);
            txtBox6.Location = new Point(229, 25);
            txtBox6.Size = new System.Drawing.Size(359, 23);
            
            this.Controls.Add(txtBox7);
            txtBox7.Location = new Point(229, 77);
            txtBox7.Size = new System.Drawing.Size(359, 23);
            
            this.groupBox1.Controls.Add(this.txtBox8);
            txtBox8.Location = new Point(80, 37);
            txtBox8.Size = new System.Drawing.Size(39, 21);
            
            this.groupBox1.Controls.Add(this.txtBox9);
            txtBox9.Location = new Point(46, 110);
            txtBox9.Size = new System.Drawing.Size(190, 21);
            txtBox9.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;

            TextBox[] arrText1 = new TextBox[] { txtBox1, txtBox2, txtBox3, txtBox4, txtBox5, txtBox6, txtBox7, txtBox8, txtBox9 };
            foreach (TextBox i in arrText1)
            {
                i.Visible = false;
                i.ReadOnly = true;
                i.BackColor = Color.Yellow;
                i.Font = new System.Drawing.Font("Arial", 9F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            }
            //ToolTip ToolTip1 = new ToolTip();
            //ToolTip1.SetToolTip(this.button3, "Удалить изображение");
         }
//==============================================================================================
//Изменение расскладки клавы в текст-боксах 
        private void i_GotFocus(object sender, System.EventArgs e) 
        {
            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(new System.Globalization.CultureInfo("uk-UA"));
        }
//==============================================================================================
//Проверка ввода в тект-боксы
        void tb_KeyPress(object sender, KeyPressEventArgs e)
        {
            TextBox t = sender as TextBox;
            string c = e.KeyChar.ToString();
            if (!Regex.Match(c, @"[а-яА-ЯґҐіІїЇєЄ'\-\e\r\b]").Success)
             {
                e.Handled = true;
                MessageBox.Show(string.Format("{0,7}", " ") + "В это поле необходимо вводить только буквы!\n\nЕсли необходимо переключите расскладку клавиатуры.",
        "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
             }
            if (e.KeyChar.Equals((char)13))
            {
                e.Handled = true;
                SendKeys.Send("{TAB}");
            }
        }
//==============================================================================================
// Первая буква в текст-боксах заглавная
        private void i_TextChanged(object sender, EventArgs e)
        {
            TextBox t = sender as TextBox;
            if (t.Text.Length == 1)
            {
                t.Text = t.Text.ToUpper();
                t.Select(t.Text.Length, 0);
            }
        }
//==============================================================================================
// Если телефон есть открываем поле для ввода номера
        private void checkBox_Tel_Click(object sender, System.EventArgs e)
        {
           if (checkBox_Tel.Checked)
            {
                maskedTextBox_Tel.Enabled = true;
                this.SelectNextControl((Control)sender, true, true, true, true);
            }
            else
            {
                maskedTextBox_Tel.Clear();
                maskedTextBox_Tel.Enabled = false;
            }
        }
//==============================================================================================
// Условия ввода в поле номер телефона
        private void maskedTextBox_Tel_KeyPress(object sender, KeyPressEventArgs e) 
        {
           if (e.KeyChar.Equals((char)13))
            {
                e.Handled = true;
                this.SelectNextControl((Control)sender, true, true, true, true);
            }
            else if (!Char.IsDigit(e.KeyChar)) 
            {
                e.Handled = true;
                MessageBox.Show("В это поле необходимо вводить только цифры!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
//==============================================================================================
//проверка номера телефона на цифры
        private void maskedTextBox_Tel_Validating(object sender, CancelEventArgs e)
        {
            if (!checkBox_Tel.Checked)
                this.SelectNextControl((Control)sender, true, true, true, true);
            else if (maskedTextBox_Tel.Text.Length != 5)
            {
                MessageBox.Show("Номер телефона состоит из 4 цифр.",
                        "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
                maskedTextBox_Tel.Focus();
            }
        }
//==============================================================================================
// Если комп есть открываем доп. контролы, если нет скрываем
        private void checkBox_Comp_Click(object sender, System.EventArgs e)
        {
            if (checkBox_Comp.Checked)
            {
                checkBox_Email.Enabled = true;
                this.SelectNextControl((Control)sender, true, true, true, true);
            }
            else
            {
                textBox_Email.Clear();
                checkBox_Email.Checked = false;
                checkBox_Email.Enabled = false;
                textBox_Email.Enabled = false;
            }
        }
//==============================================================================================
// Если почта есть открываем поле для её ввода
        private void checkBox_Email_Click(object sender, System.EventArgs e)
        {
            if (checkBox_Email.Checked)
            {
                textBox_Email.Enabled = true;
                this.SelectNextControl((Control)sender, true, true, true, true);
            }
            else
            {
                textBox_Email.Clear();
                textBox_Email.Enabled = false;
            }
        }
//==============================================================================================
// Меняем для ввода расскладку в поле адресс почты 
        private void textBox_Email_GotFocus(object sender, System.EventArgs e) 
        {
            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(new System.Globalization.CultureInfo("en-US"));
        }
//==============================================================================================
// Условия ввода адресса почты
        void textBox_Email_KeyPress(object sender, KeyPressEventArgs e)
        {
            string c = e.KeyChar.ToString();
            if (!Regex.Match(c, @"[a-zA-Z0-9_\-\e\r\b\.]").Success)
            {
                e.Handled = true;
                MessageBox.Show("Только латиница, цифры, тире и нижнее подчеркивание!",
        "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            if (e.KeyChar.Equals((char)13))
            {
                e.Handled = true;
                SendKeys.Send("{TAB}");
            }
        }
//==============================================================================================
// Поиск фотографии и ввод её в пикча-бокс
        private void button_Foto_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = @"d:\Mydocuments\Изображения\лица из банка\";
            openFileDialog1.Filter = "тип JPG или BMP|*.jpg;*.jpeg;*.bmp";
            openFileDialog1.Title = "Поиск фотографии";
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                pictureBox1.Image = Image.FromFile(openFileDialog1.FileName);
                GlobalData.pictureBox = pictureBox1.Image;
                button3.Enabled = true;
             }
        }
//==============================================================================================   
// Кнопка ОК в двух вариантах
        private void button1_Click(object sender, EventArgs e)
        {
            if (p2 == 1)
            {
                   GlobalData.querySQL = "UPDATE Table1 SET Tab='" + txtBox1.Text + "',F='"+ txtBox2.Text + "',I='"  + txtBox3.Text + "',O='" + txtBox4.Text + 
                    "',Bithday='" + dateTimePicker1.Value.ToShortDateString() + "',Otdel='" + txtBox6.Text +"',Dolzhn='" + txtBox7.Text + "',Tel='" + txtBox8.Text +
                    "',NameComp='" + Comp + "',Email='" + email + "',Foto=@MyImg WHERE id=" + GlobalData.id;
                GlobalData.prizn = 3;
                ActiveForm.Close();
            }
            else if (((textBox_Tab.Text == "") || (textBox_F.Text == "") || (textBox_I.Text == "") || (textBox_O.Text == "") || (comboBox_Otdel.Text == "") || (comboBox_Dolzhn.Text == "") || !dateTimePicker1.Checked))
            {
                MessageBox.Show(string.Format("{0,15}", " ") + "Не все поля заполнены!\n\n" +
            "Поля обязательные для ввода помечены (*)", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else if ((checkBox_Email.Checked) && (textBox_Email.Text == ""))
            { 
                MessageBox.Show(string.Format("{0,5}", " ") + "Введите адрес электронной почты или\n\nоткажитесь от неё убрав чек напротив Email.",
                    "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
            }
            else
            {
                p1 = 1;
                button_Foto.Enabled = false;
                
                textBox_Tab.Visible = false;
                txtBox1.Visible = true;
                txtBox1.Text = textBox_Tab.Text.Trim();
                
                textBox_F.Visible = false;
                txtBox2.Visible = true;
                txtBox2.Text = textBox_F.Text.Trim();
               
                textBox_I.Visible = false;
                txtBox3.Visible = true;
                txtBox3.Text = textBox_I.Text.Trim();
                
                textBox_O.Visible = false;
                txtBox4.Visible = true;
                txtBox4.Text = textBox_O.Text.Trim();
                
                dateTimePicker1.Visible = false;
                txtBox5.Visible = true;
                txtBox5.Text = dateTimePicker1.Text;
                
                comboBox_Otdel.Visible = false;
                txtBox6.Visible = true;
                txtBox6.Text = comboBox_Otdel.Text.Trim();
               
                comboBox_Dolzhn.Visible = false;
                txtBox7.Visible = true;
                txtBox7.Text = comboBox_Dolzhn.Text.Trim();
                
                if (checkBox_Tel.Checked)
                {
                    checkBox_Tel.Visible = false;
                    checkBox_Tel.BackColor = Color.Yellow;
                    maskedTextBox_Tel.Visible = false;
                    txtBox8.Visible = true;
                    txtBox8.Text = maskedTextBox_Tel.Text;
                 }
                else
                {
                    label7.Visible = false;
                    checkBox_Tel.Visible = false;
                    maskedTextBox_Tel.Visible = false;
                    txtBox8.Text = " ";
                }

                if (checkBox_Comp.Checked)
                {
                    checkBox_Comp.Visible = false;
                }
                else
                {
                    checkBox_Comp.Visible = false;
                    label12.Visible = false;
                    label9.Visible = false;
                    textBox_Email.Visible = false;
                    label10.Visible = false;
                    checkBox_Email.Visible = false;
                }

                if (checkBox_Email.Checked)
                {
                    checkBox_Email.Visible = false;
                    textBox_Email.Visible = false;
                    txtBox9.Visible = true;
                    txtBox9.Text = textBox_Email.Text.Trim();
                    email = txtBox9.Text + label10.Text;
                }
                else
                {
                    checkBox_Email.Visible = false;
                    label9.Visible = false;
                    textBox_Email.Visible = false;
                    label10.Visible = false;
                    email = " ";
                }
                p2 = 1;
                 if (checkBox_Comp.Checked)
                    switch (comboBox_Otdel.SelectedIndex)
                    {
                        case 0:
                            Comp = "C9U-101-" + textBox_Tab.Text;
                            break;
                        case 1:
                            Comp = "C9U-102-" + textBox_Tab.Text;
                            break;
                        case 2:
                            Comp = "C9U-103-" + textBox_Tab.Text;
                            break;
                        case 3:
                            Comp = "C9U-104-" + textBox_Tab.Text;
                            break;
                        case 4:
                            Comp = "C9U-106-" + textBox_Tab.Text;
                            break;
                        case 5:
                            Comp = "C9U-108-" + textBox_Tab.Text;
                            break;
                        case 6:
                            Comp = "C9U-109-" + textBox_Tab.Text;
                            break;
                        case 7:
                            Comp = "C9U-112-" + textBox_Tab.Text;
                            break;
                        case 8:
                            Comp = "C9U-115-" + textBox_Tab.Text;
                            break;
                        case 9:
                            Comp = "C9U-207-" + textBox_Tab.Text;
                            break;
                        case 10:
                            Comp = "C9U-208-" + textBox_Tab.Text;
                            break;
                        case 11:
                            Comp = "C9U-209-" + textBox_Tab.Text;
                            break;
                        case 12:
                            Comp = "C9U-301-" + textBox_Tab.Text;
                            break;
                        case 13:
                            Comp = "C9U-401-" + textBox_Tab.Text;
                            break;
                        case 14:
                            Comp = "C9U-901-" + textBox_Tab.Text;
                            break;
                    }
                else
                      Comp = " ";
              }
        }
//==============================================================================================
// кнопка удаления фотографии
        private void button3_Click(object sender, EventArgs e)
        {
            pictureBox1.Image = null;
            pictureBox1.Invalidate();
            button3.Enabled = false;
        }
    }
}