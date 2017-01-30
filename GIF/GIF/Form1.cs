using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Drawing.Imaging;

namespace WindowsFormsApplication2
{
    public partial class Form1 : Form
    {
        string FilesGif, Catalog;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.TextLength == 0)
                MessageBox.Show("Найдите файл GIF!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            else if (!File.Exists(FilesGif))
                MessageBox.Show("Такого файла GIF не существует!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else if (textBox2.TextLength == 0)
                MessageBox.Show("Определите директорию для кадров!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            else if (!Directory.Exists(Catalog))
                MessageBox.Show("Такой директории не существует!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                using (Image gifImage = Image.FromFile(@FilesGif))
                {
                    FrameDimension dimension = new FrameDimension(gifImage.FrameDimensionsList[0]);
                    progressBar1.Minimum = 0;
                    progressBar1.Maximum = gifImage.GetFrameCount(dimension) - 1;
                    progressBar1.Step = 1;
                    for (int index = 0; index < gifImage.GetFrameCount(dimension) - 1; index++)
                    {
                        gifImage.SelectActiveFrame(dimension, index);
                        using (Bitmap tmpBmp = new Bitmap(gifImage.Width, gifImage.Height, PixelFormat.Format32bppPArgb))
                        using (Graphics g = Graphics.FromImage(tmpBmp))
                        {
                            g.DrawImage((Image)gifImage, 0, 0);
                            tmpBmp.Save(string.Format(@Catalog + "\\{0:000}.png", index));
                        }
                        progressBar1.PerformStep();
                    }
                }
                Text = new DirectoryInfo(@Catalog).GetFiles().Length.ToString();
                MessageBox.Show("Количество файлов (кадров) в папке " + Text);
                progressBar1.Value = 0;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = @"d:\Mydocuments\Изображения\";
            openFileDialog1.Filter = "GIF files (*.gif)|*.gif";
            openFileDialog1.Title = "Поиск GIF файла";
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FilesGif = openFileDialog1.FileName;
                    textBox1.Text = FilesGif;
                    textBox1.Visible = true;                   
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Файл GIF не может быть прочитан: " + ex.Message);
                }
             }
           
           }

        private void button4_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                Catalog = folderBrowserDialog1.SelectedPath;
                textBox2.Text = Catalog + "\\";
                textBox2.Visible = true;
            }
        }

    }
}