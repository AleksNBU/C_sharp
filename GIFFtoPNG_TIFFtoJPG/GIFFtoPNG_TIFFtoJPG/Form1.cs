using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace WindowsFormsApplication2
{
    public partial class Form1 : Form
    {
        string FilesGif, Catalog, FilesTiff, FilesJpg;
        public Form1()
        {
            InitializeComponent();
        }
//=======================================================================================================================================
//             GIF to PNG
//=======================================================================================================================================
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
                    progressBar1.Maximum = gifImage.GetFrameCount(dimension);
                    progressBar1.Step = 1;
                    try
                    {
                        for (int index = 0; index < gifImage.GetFrameCount(dimension); index++)
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
                    catch (Exception ex)
                    {
                        MessageBox.Show("Вовремя конвертации произошла ошибка: " + ex.Message);
                    }
                }
                MessageBox.Show("Количество файлов (кадров) PNG - " + progressBar1.Maximum.ToString());
                progressBar1.Value = 0;
                textBox3.Clear();
                textBox2.Clear();
                checkBox1.Checked = false;
                button4.Enabled = false;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "GIF files (*.gif)|*.gif";
            openFileDialog1.Title = "Поиск GIF файла";
            openFileDialog1.Multiselect = false;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FilesGif = openFileDialog1.FileName;
                    textBox3.Text = FilesGif;
                    checkBox1.Enabled = true;
                    button4.Enabled = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Файл GIF не может быть прочитан: " + ex.Message);
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                Catalog = folderBrowserDialog1.SelectedPath;
                textBox2.Text = Catalog + "\\";
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {

                button4.Enabled = false;
                Catalog = Path.GetDirectoryName(FilesGif);
                textBox2.Text = Catalog + "\\";
            }
            else
            {
                button4.Enabled = true;
                textBox2.Clear();
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (textBox2.Text == String.Empty)
                button1.Enabled = false;
            else button1.Enabled = true;
        }

//====================================================================================================================
//                                 TIFF to JPG
//====================================================================================================================
        private void button5_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "TIFF files (.tif, .tiff)|*.tif;*.tiff";
            openFileDialog1.InitialDirectory = @"d:\Mydocuments\Изображения\";
            openFileDialog1.Title = "Поиск TIFF файла";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FilesTiff = openFileDialog1.FileName;
                    textBox6.Text = FilesTiff;
                    checkBox2.Enabled = true;
                    button2.Enabled = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Файл TIFF не может быть прочитан: " + ex.Message, "Error");
                }
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                Catalog = Path.GetDirectoryName(FilesTiff);
                textBox5.Text = Catalog + "\\";
            }
            else
            {
                button6.Enabled = false;
                textBox5.Clear();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                Catalog = folderBrowserDialog1.SelectedPath;
                textBox5.Text = Catalog + "\\";
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                 using (Image imageFile = Image.FromFile(FilesTiff))
                {
                    FrameDimension frameDimensions = new FrameDimension(
                        imageFile.FrameDimensionsList[0]);
                    int frameNum = imageFile.GetFrameCount(frameDimensions);
                    string[] jpegPaths = new string[frameNum];

                    progressBar2.Minimum = 0;
                    progressBar2.Maximum = frameNum;
                    progressBar2.Step = 1;

                    for (int frame = 0; frame < frameNum; frame++)
                    {
                        // Selects one frame at a time and save as jpeg.
                        imageFile.SelectActiveFrame(frameDimensions, frame);
                        using (Bitmap bmp = new Bitmap(imageFile))
                        {
                            jpegPaths[frame] = String.Format("{0}\\{1}{2}.jpg", @Catalog, Path.GetFileNameWithoutExtension(FilesTiff), frame);
                            bmp.Save(jpegPaths[frame], ImageFormat.Jpeg);
                        }
                        progressBar2.PerformStep();
                    }
                    MessageBox.Show("Количество изображений в файле TIFF - " + frameNum.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Вовремя конвертации произошла ошибка: " + ex.Message, "Error");
            }
            progressBar2.Value = 0;
            textBox6.Clear();
            textBox5.Clear();
            checkBox2.Checked = false;
            button2.Enabled = false;
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            if (textBox5.Text == String.Empty)
                button6.Enabled = false;
            else button6.Enabled = true;
        }
//====================================================================================================================
//                                 JPG to TIFF
//====================================================================================================================      
        private void button9_Click(object sender, EventArgs e)
        {
            openFileDialog1.Multiselect = true;
            openFileDialog1.Filter = "JPG files (.jpg, .jpeg)|*.jpg;*.jpeg";
            openFileDialog1.Title = "Поиск JPG файла (файлов)";
           if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string nameFiles = "";
                    try
                    {
                        if (openFileDialog1.FileNames.Count() > 1)
                        {
                            foreach (String file in openFileDialog1.FileNames)
                            {
                                nameFiles += Path.GetFileName(file) + ", ";
                            }
                            nameFiles = nameFiles.Remove(nameFiles.LastIndexOf(","));
                            FilesJpg = Path.GetDirectoryName(openFileDialog1.FileName);
                            textBox8.Text = FilesJpg + "\\ (" + nameFiles + ")";

                        }
                        else
                        {
                            FilesJpg = openFileDialog1.FileName;
                            textBox8.Text = FilesJpg;
                            FilesJpg = Path.GetDirectoryName(FilesJpg);
                        }
                        checkBox4.Enabled = true;
                        checkBox3.Enabled = true;
                        button8.Enabled = true;
                        
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Файл TIFF не может быть прочитан: " + ex.Message, "Error");
                    }
            }
       }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked)
            {
                button8.Enabled = false;
                Catalog = FilesJpg;
                textBox7.Text = FilesJpg + "\\";
            }
            else
            {
                button8.Enabled = true;
                textBox7.Clear();
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                Catalog = folderBrowserDialog1.SelectedPath;
                textBox7.Text = Catalog + "\\";
            }
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            if (textBox7.Text == String.Empty)
                button7.Enabled = false;
            else button7.Enabled = true;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            ConvertJpegToTiff(openFileDialog1.FileNames, checkBox4.Checked, Catalog);
            MessageBox.Show("Image conversion completed.");
        }

        public static string[] ConvertJpegToTiff(string[] fileNames, bool isMultipage, string toTiff)
        {
            EncoderParameters encoderParams = new EncoderParameters(1);
            ImageCodecInfo tiffCodecInfo = ImageCodecInfo.GetImageEncoders()
                .First(ie => ie.MimeType == "image/tiff");

            string[] tiffPaths = null;
            if (isMultipage)
            {
                tiffPaths = new string[1];
                Image tiffImg = null;
                try
                {
                    for (int i = 0; i < fileNames.Length; i++)
                    {
                        if (i == 0)
                        {
                            tiffPaths[i] = String.Format("{0}\\{1}.tif", Path.GetDirectoryName(fileNames[i]),
                                Path.GetFileNameWithoutExtension(fileNames[i]));

                            // Initialize the first frame of multipage tiff.
                            tiffImg = Image.FromFile(fileNames[i]);
                            encoderParams.Param[0] = new EncoderParameter(
                                System.Drawing.Imaging.Encoder.SaveFlag, (long)EncoderValue.MultiFrame);
                            tiffImg.Save(tiffPaths[i], tiffCodecInfo, encoderParams);
                        }
                        else
                        {
                            // Add additional frames.
                            encoderParams.Param[0] = new EncoderParameter(
                                System.Drawing.Imaging.Encoder.SaveFlag, (long)EncoderValue.FrameDimensionPage);
                            using (Image frame = Image.FromFile(fileNames[i]))
                            {
                                tiffImg.SaveAdd(frame, encoderParams);
                            }
                        }

                        if (i == fileNames.Length - 1)
                        {
                            // When it is the last frame, flush the resources and closing.
                            encoderParams.Param[0] = new EncoderParameter(
                                System.Drawing.Imaging.Encoder.SaveFlag, (long)EncoderValue.Flush);
                            tiffImg.SaveAdd(encoderParams);
                        }
                    }
                }
                finally
                {
                    if (tiffImg != null)
                    {
                        tiffImg.Dispose();
                        tiffImg = null;
                    }
                }
            }
            else
            {
                tiffPaths = new string[fileNames.Length];

                for (int i = 0; i < fileNames.Length; i++)
                {
                    tiffPaths[i] = String.Format("{0}\\{1}.tif",
                        Path.GetDirectoryName(fileNames[i]),
                        Path.GetFileNameWithoutExtension(fileNames[i]));

                    // Save as individual tiff files.
                    using (Image tiffImg = Image.FromFile(fileNames[i]))
                    {
                        tiffImg.Save(tiffPaths[i], ImageFormat.Tiff);
                    }
                }
            }

            return tiffPaths;
        }
    }
}