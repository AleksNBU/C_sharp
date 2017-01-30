//==============================================================================================
// Форма удаления сведений о сотруднике
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
    public partial class Form_Delete : Form
    {
        private TextBox txtBox1 = new TextBox();
        private TextBox txtBox2 = new TextBox();
        private TextBox txtBox3 = new TextBox();
        private TextBox txtBox4 = new TextBox();
        private TextBox txtBox5 = new TextBox();
        private TextBox txtBox6 = new TextBox();
        private TextBox txtBox7 = new TextBox();
        private TextBox txtBox8 = new TextBox();
        private TextBox txtBox9 = new TextBox();
        
        public Form_Delete()
        {
            InitializeComponent();
        }
//==============================================================================================
// Создание и инициализация новых контролов, в которых будет инфа об удаляемом сотруднике
        private void Form2_Load(object sender, EventArgs e)
        {
            // Новые контролы
            TextBox[] arrText = new TextBox[] { txtBox1, txtBox2, txtBox3, txtBox4, txtBox5, txtBox6, txtBox7, txtBox8, txtBox9 };
            foreach (TextBox i in arrText)
            {
                i.ReadOnly = true;
                i.BackColor = Color.Yellow;
                i.Font = new System.Drawing.Font("Arial", 9F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            }
            this.Controls.Add(txtBox1);
            txtBox1.Text = GlobalData.Tab;
            txtBox1.Location = new Point(11, 70);
            txtBox1.Size = new System.Drawing.Size(35, 20);
            
            this.Controls.Add(txtBox2);
            txtBox2.Text = GlobalData.F;
            txtBox2.Location = new Point(11, 121);
            txtBox2.Size = new System.Drawing.Size(190, 21);
            
            this.Controls.Add(txtBox3);
            txtBox3.Text = GlobalData.I;
            txtBox3.Location = new Point(11, 175);
            txtBox3.Size = new System.Drawing.Size(190, 21);
           
            this.Controls.Add(txtBox4);
            txtBox4.Text = GlobalData.O;
            txtBox4.Location = new Point(11, 229);
            txtBox4.Size = new System.Drawing.Size(190, 21);
           
            this.Controls.Add(txtBox5);
            txtBox5.Text = GlobalData.Bithday.ToLongDateString();
            txtBox5.Location = new Point(11, 285);
            txtBox5.Size = new System.Drawing.Size(120, 21);
            
            this.Controls.Add(txtBox6);
            txtBox6.Text = GlobalData.Otdel;
            txtBox6.Location = new Point(229, 70);
            txtBox6.Size = new System.Drawing.Size(359, 23);
            
            this.Controls.Add(txtBox7);
            txtBox7.Text = GlobalData.Dolzhn;
            txtBox7.Location = new Point(229, 121);
            txtBox7.Size = new System.Drawing.Size(359, 23);

            if (GlobalData.Tel.Trim().Length != 0)
            {
                this.groupBox1.Controls.Add(this.txtBox8);
                txtBox8.Text = GlobalData.Tel;
                txtBox8.Location = new Point(80, 37);
                txtBox8.Size = new System.Drawing.Size(39, 21);
            }
            else
                label7.Visible = false;

            if (GlobalData.Comp.Trim().Length == 0)
                label12.Visible = false;

            if (GlobalData.Email != " ")
            {
                this.groupBox1.Controls.Add(this.txtBox9);
                txtBox9.Text = GlobalData.Email;
                txtBox9.Location = new Point(45, 110);
                txtBox9.Size = new System.Drawing.Size(300, 21);
                txtBox9.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            }
            else
                label9.Visible = false;

            pictureBox1.Image = GlobalData.pictureBox;
        }
//==============================================================================================
// Кнопка Отмена
        private void button2_Click(object sender, EventArgs e)
        {
            GlobalData.prizn = 4;
            this.Close();
        }
//==============================================================================================
// Кнопка ОК, согласие на удаление
        private void button1_Click(object sender, EventArgs e)
        {
            GlobalData.prizn = 2;
            ActiveForm.Close();
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
    }
}
