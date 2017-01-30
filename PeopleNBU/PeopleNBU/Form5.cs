//==============================================================================================
// Форма предоставления полной информации о сотруднике
//==============================================================================================
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PeopleNBU
{
    public partial class Form5 : Form
    {
        private TextBox txtBox1 = new TextBox();//Таб. номер
        private TextBox txtBox2 = new TextBox();//Фамилия
        private TextBox txtBox3 = new TextBox();//Имя
        private TextBox txtBox4 = new TextBox();//Отчество
        private TextBox txtBox5 = new TextBox();//Дата рождения 
        private TextBox txtBox6 = new TextBox();//Отдел  
        private TextBox txtBox7 = new TextBox();//Должность  
        private TextBox txtBox8 = new TextBox();//Телефон
        private TextBox txtBox9 = new TextBox();//Email
        private TextBox txtBox10 = new TextBox();//Имя компа
        private TextBox txtBox11 = new TextBox();//Возраст
        private TextBox txtBox12 = new TextBox();//Прописка
        private TextBox txtBox13 = new TextBox();//Место рождения
        private TextBox txtBox14 = new TextBox();//Пасспорт
        private TextBox txtBox15 = new TextBox();//Идентификационный номер
              
        public Form5()
        {
            InitializeComponent();
        }
//==============================================================================================
// Инициализация новых контролов, в которых будет инфа о сотруднике 
        private void Form5_Load(object sender, EventArgs e)
        {
            button1.Left = (this.ClientSize.Width - button1.Width) / 2;
            button1.Top = button1.Top - 159;
            this.Height = this.Height - 159;
            // Новые контролы
            TextBox[] arrText = new TextBox[] { txtBox1, txtBox2, txtBox3, txtBox4, txtBox5, txtBox6, txtBox7, txtBox8, txtBox9, txtBox10, txtBox11, txtBox12, txtBox13, txtBox14, txtBox15 };
            foreach (TextBox i in arrText)
            {
                i.ReadOnly = true;
                i.BackColor = Color.Yellow;
                i.Font = new System.Drawing.Font("Arial", 9F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            }
            this.Controls.Add(txtBox1);
            txtBox1.Text = GlobalData.Tab;
            txtBox1.Location = new Point(11, 38);
            txtBox1.Size = new System.Drawing.Size(35, 21);
            
            this.Controls.Add(txtBox2);
            txtBox2.Text = GlobalData.F;
            txtBox2.Location = new Point(11, 89);
            txtBox2.Size = new System.Drawing.Size(190, 21);
            
            this.Controls.Add(txtBox3);
            txtBox3.Text = GlobalData.I;
            txtBox3.Location = new Point(11, 143);
            txtBox3.Size = new System.Drawing.Size(190, 21);
           
            this.Controls.Add(txtBox4);
            txtBox4.Text = GlobalData.O;
            txtBox4.Location = new Point(11, 197);
            txtBox4.Size = new System.Drawing.Size(190, 21);
           
            this.Controls.Add(txtBox5);
            txtBox5.Text = GlobalData.Bithday.ToLongDateString();
            txtBox5.Location = new Point(11, 253);
            txtBox5.Size = new System.Drawing.Size(120, 21);

            this.Controls.Add(txtBox11);
            txtBox11.Text = CalculateAge(GlobalData.Bithday).ToString();
            txtBox11.Location = new Point(174, 253);
            txtBox11.Size = new System.Drawing.Size(23, 21);
            
            this.Controls.Add(txtBox6);
            txtBox6.Text = GlobalData.Otdel;
            txtBox6.Location = new Point(229, 38);
            txtBox6.Size = new System.Drawing.Size(359, 23);
            
            this.Controls.Add(txtBox7);
            txtBox7.Text = GlobalData.Dolzhn;
            txtBox7.Location = new Point(229, 89);
            txtBox7.Size = new System.Drawing.Size(359, 23);

            if (GlobalData.Tel.Trim().Length != 0)
            {
                this.groupBox1.Controls.Add(this.txtBox8);
                txtBox8.Text = GlobalData.Tel;
                txtBox8.Location = new Point(75, 37);
                txtBox8.Size = new System.Drawing.Size(39, 21);
            }
            else
                label7.Visible = false;

            if (GlobalData.Comp.Trim().Length == 0)
                label12.Visible = false;
            else
            {
                this.groupBox1.Controls.Add(this.txtBox10);
                txtBox10.Text = GlobalData.Comp;
                txtBox10.Location = new Point(118, 75);
                txtBox10.Size = new System.Drawing.Size(90, 21);
            }

            if (GlobalData.Email.Trim().Length != 0)
            {
                this.groupBox1.Controls.Add(this.txtBox9);
                txtBox9.Text = GlobalData.Email;
                txtBox9.Location = new Point(45, 110);
                txtBox9.Size = new System.Drawing.Size(300, 21);
                txtBox9.TextAlign = System.Windows.Forms.HorizontalAlignment.Left;
            }
            else
                label9.Visible = false;
                        
            pictureBox1.Image = GlobalData.pictureBox;
        }
//==============================================================================================
// Кнопка Отмена
        private void button2_Click(object sender, EventArgs e)
        {
            ActiveForm.Close();
        }
//==============================================================================================
// Кнопка ОК
        private void button1_Click(object sender, EventArgs e)
        {
            GlobalData.prizn = 4;
            ActiveForm.Close();
        }
//==============================================================================================
// Если отстутствует фотография
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
// вычисление возраста сотрудника
        public static int CalculateAge(DateTime BirthDate)
        {
            int YearsPassed = DateTime.Now.Year - BirthDate.Year;
            if (DateTime.Now.Month < BirthDate.Month || (DateTime.Now.Month == BirthDate.Month && DateTime.Now.Day < BirthDate.Day))
            {
                YearsPassed--;
            }
            return YearsPassed;
        }
//==============================================================================================
// Раскрытие дополнительной информации
        int bt = 0;
        private void button2_Click_1(object sender, EventArgs e)
        {
            if (bt == 0)
            {
                textBox1.Visible = true;
                bt = 1;
            }
            else if ((bt == 1) & (textBox1.Text == "qwe123"))
            {
                this.Height = this.Height + 159;
                this.button2.Top = button2.Top + 159;
                this.button2.Image = global::PeopleNBU.Properties.Resources.arrow_up;
                textBox1.Visible = false;
                bt = 2;
                button1.Top = button1.Top + 159;
                groupBox3.Visible = true;

                this.groupBox3.Controls.Add(txtBox12);
                txtBox12.Text = GlobalData.Adress;
                txtBox12.Location = new Point(6, 52);
                txtBox12.Size = new Size(590, 21);

                this.groupBox3.Controls.Add(txtBox13);
                txtBox13.Text = GlobalData.BIRTHPLACE;
                txtBox13.Location = new Point(6, 102);
                txtBox13.Size = new Size(590, 21);

                this.groupBox3.Controls.Add(txtBox14);
                txtBox14.Text = GlobalData.Passport;
                txtBox14.Location = new Point(622, 52);
                txtBox14.Size = new Size(80, 21);

                this.groupBox3.Controls.Add(txtBox15);
                txtBox15.Text = GlobalData.IPN;
                txtBox15.Location = new Point(622, 102);
                txtBox15.Size = new Size(80, 21);
            }
            else if (bt == 2)
            {
                groupBox3.Visible = false;
                this.button1.Top = button1.Top - 159;
                this.Height = this.Height - 159;
                this.button2.Top = button2.Top - 159;
                textBox1.Text = "";
                textBox1.Visible = false;
                bt = 0;
                this.button2.Image = global::PeopleNBU.Properties.Resources.arrow_down;
            }
            else if ((bt == 1) & (textBox1.Text == ""))
            {
                textBox1.Visible = false;
                bt = 0;
            }
        }
    }
}
