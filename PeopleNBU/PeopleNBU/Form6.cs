//==============================================================================================
// Форма сохранения в файл сведений о сотруднике
//============================================================================================== 
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel=Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.Reflection;
using System.IO;

namespace PeopleNBU
{
    public partial class Form6 : Form
    {
        string filter, fileName;
        int ii, jj;
        public Form6()
        {
            InitializeComponent();
        }
//==============================================================================================
// Диалоговое окно сохранения файла
        private void button1_Click(object sender, EventArgs e)
        {
            Form1 main = this.Owner as Form1;
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.InitialDirectory = @"D:\";
            saveFileDialog1.Filter = filter;
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileName = saveFileDialog1.FileName;

                main.Refresh();
                this.Refresh();
                if ((radioButton3.Checked))
                {
                    Excel.Application xlApp = null;
                    Excel.Workbook xlWorkBook = null;
                    Excel.Worksheet xlWorkSheet = null;
                    object Value = System.Reflection.Missing.Value;
                    xlApp = new Excel.ApplicationClass();
                    xlApp.DisplayAlerts = false;
                    xlApp.Visible = false;
                    xlWorkBook = xlApp.Workbooks.Add(Value);
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                    xlWorkSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
                    xlWorkSheet.PageSetup.LeftMargin = 10;
                    xlWorkSheet.PageSetup.RightMargin = 10;
                    xlWorkSheet.PageSetup.TopMargin = 10;
                    xlWorkSheet.PageSetup.BottomMargin = 10;
                    xlWorkSheet.PageSetup.Zoom = 75;
                    xlWorkSheet.Range["A2:J2"].MergeCells = true;
                    xlWorkSheet.Range["A2:J2"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    xlWorkSheet.Range["A2:J2"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    xlWorkSheet.Range["A2:J2"].RowHeight = 40;
                    xlWorkSheet.Range["A2:J2"].Font.Size = 16;
                    xlWorkSheet.Range["A2:J2"].Value = "Реестр сотрудников\nНационального банка Украины в Харьковской области";
                    xlWorkSheet.Range["A4"].Value = "Табельный\nномер";
                    xlWorkSheet.Range["B4"].Value = "Фамилия";
                    xlWorkSheet.Range["C4"].Value = "Имя";
                    xlWorkSheet.Range["D4"].Value = "Отчество";
                    xlWorkSheet.Range["E4"].Value = "Дата\nрождения";
                    xlWorkSheet.Range["F4"].Value = "Отдел";
                    xlWorkSheet.Range["G4"].Value = "Должность";
                    xlWorkSheet.Range["H4"].Value = "Телефон";
                    xlWorkSheet.Range["I4"].Value = "Имя\nкомпьютера";
                    xlWorkSheet.Range["J4"].Value = "Email";
                    xlWorkSheet.Range["A4:J4"].HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    xlWorkSheet.Range["A4:J4"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    xlWorkSheet.Range["A4:J4"].Borders.LineStyle = Excel.XlBordersIndex.xlEdgeLeft;
                    xlWorkSheet.Range["A:A"].HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    xlWorkSheet.Range["A4:J4"].RowHeight = 30;
                    xlWorkSheet.Range["H:H", "I:I"].HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    xlWorkSheet.Range["A:A"].NumberFormat = "@";
                    xlWorkSheet.Range["A:A"].ColumnWidth = 11;
                    if ((radioButton1.Checked))
                    {
                        xlWorkSheet.Range["A6"].Value = GlobalData.Tab;
                        xlWorkSheet.Range["B6"].Value = GlobalData.F;
                        xlWorkSheet.Range["C6"].Value = GlobalData.I;
                        xlWorkSheet.Range["D6"].Value = GlobalData.O;
                        xlWorkSheet.Range["E6"].Value = GlobalData.Bithday.ToShortDateString();
                        xlWorkSheet.Range["F6"].Value = GlobalData.Otdel;
                        xlWorkSheet.Range["G6"].Value = GlobalData.Dolzhn;
                        xlWorkSheet.Range["H6"].Value = GlobalData.Tel;
                        xlWorkSheet.Range["I6"].Value = GlobalData.Comp;
                        xlWorkSheet.Range["J6"].Value = GlobalData.Email;
                        xlWorkSheet.Range["A:J"].Columns.AutoFit();
                        xlWorkSheet.Range["A5:A7"].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Range["J5:J7"].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Range["A7:J7"].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    }
                    else if ((radioButton2.Checked))
                    {
                        button2.Visible = false;
                        int y;
                        progressBar1.Visible = true;
                        progressBar1.Minimum = 0;
                        progressBar1.Maximum = main.dataGridView1.RowCount;
                        progressBar1.Step = 1;
                        for (int i = 0; i < main.dataGridView1.RowCount; i++)
                        {
                            for (int j = 0; j < 10; j++)
                            {
                                try
                                {
                                    xlWorkSheet.Cells[i + 6, j + 1] = main.dataGridView1[j + 1, i].Value;
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show("Невозможно сформировать файл. " + ex.Message);
                                }
                                if (j == 0)
                                {
                                    Excel.Range rng = (Excel.Range)xlWorkSheet.Cells[i + 6, j + 1];
                                    rng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                                }
                                if (j == 9)
                                {
                                    Excel.Range rng = (Excel.Range)xlWorkSheet.Cells[i + 6, j + 1];
                                    rng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                                }
                                jj = j;
                            }
                            ii = i;
                            progressBar1.PerformStep();
                            y = 100 * i / main.dataGridView1.RowCount + 1;
                            label1.Text = y.ToString() + " %";
                        }
                        xlWorkSheet.Range["A:J"].Columns.AutoFit();
                        xlWorkSheet.Range["A5"].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Range["J5"].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                        Excel.Range rng1 = (Excel.Range)xlWorkSheet.get_Range(xlWorkSheet.Cells[ii + 6, 1], xlWorkSheet.Cells[ii + 6, jj + 1]);
                        rng1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                     }
                    xlWorkBook.SaveAs(fileName, Excel.XlFileFormat.xlWorkbookNormal, Missing.Value, Missing.Value, Missing.Value,
                        Missing.Value, Excel.XlSaveAsAccessMode.xlExclusive, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                    xlWorkBook.Close(false, Value, Value);
                    xlApp.Quit();
                    releaseObject(xlWorkSheet);
                    releaseObject(xlWorkBook);
                    releaseObject(xlApp);
                }
                else if ((radioButton4.Checked))
                {
                    object oMissing = System.Reflection.Missing.Value;
                    object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */
                    Word._Application oWord;
                    Word._Document oDoc;
                    oWord = new Word.Application();
                    oWord.Visible = false;
                    oDoc = oWord.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                    oWord.Selection.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
                    oWord.Selection.ParagraphFormat.SpaceAfter = 0;
                    oWord.Selection.PageSetup.RightMargin = 10;
                    oWord.Selection.PageSetup.TopMargin = 5;
                    oWord.Selection.PageSetup.LeftMargin = 10;
                    oWord.Selection.PageSetup.BottomMargin = 5;
                    Word.Range oPara = oDoc.Range(ref oMissing, ref oMissing);
                    oPara.Font.Size = 9;
                    oPara.Text = "Реестр сотрудников\n";
                    oPara.InsertAfter("Национального банка Украины в Харьковской области");
                    oPara.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    Word.Table oTable;
                    Word.Range rng2 = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    if ((radioButton1.Checked))
                        oTable = oDoc.Tables.Add(rng2, 3, 10, ref oMissing, ref oMissing);
                    else
                        oTable = oDoc.Tables.Add(rng2, main.dataGridView1.RowCount + 2, 10, ref oMissing, ref oMissing);
                    oTable.Range.Font.Size = 6;
                    oTable.Cell(1, 1).Range.Text = "Табельный\nномер";
                    oTable.Cell(1, 2).Range.Text = "Фамилия";
                    oTable.Cell(1, 3).Range.Text = "Имя";
                    oTable.Cell(1, 4).Range.Text = "Отчество";
                    oTable.Cell(1, 5).Range.Text = "Дата\nрождения";
                    oTable.Cell(1, 6).Range.Text = "Отдел";
                    oTable.Cell(1, 7).Range.Text = "Должность";
                    oTable.Cell(1, 8).Range.Text = "Телефон";
                    oTable.Cell(1, 9).Range.Text = "Имя\nкомпьютера";
                    oTable.Cell(1, 10).Range.Text = "Email";
                    for (int i = 1; i < 11; i++)
                    {
                        oTable.Cell(1, i).Range.Font.Bold = 1;
                        oTable.Cell(1, i).Range.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    }
                    Word.Range rows = oTable.Rows[1].Range;
                    rows.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    rows.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    rows.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    oTable.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;
                    oTable.Cell(2, 1).Range.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    oTable.Cell(2, 10).Range.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    if ((radioButton1.Checked))
                    {
                        oTable.Cell(3, 1).Range.Text = GlobalData.Tab;
                        oTable.Cell(3, 2).Range.Text = GlobalData.F;
                        oTable.Cell(3, 3).Range.Text = GlobalData.I;
                        oTable.Cell(3, 4).Range.Text = GlobalData.O;
                        oTable.Cell(3, 5).Range.Text = GlobalData.Bithday.ToShortDateString();
                        oTable.Cell(3, 6).Range.Text = GlobalData.Otdel;
                        oTable.Cell(3, 7).Range.Text = GlobalData.Dolzhn;
                        oTable.Cell(3, 8).Range.Text = GlobalData.Tel;
                        oTable.Cell(3, 9).Range.Text = GlobalData.Comp;
                        oTable.Cell(3, 10).Range.Text = GlobalData.Email;
                        oTable.Cell(3, 1).Range.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                        oTable.Cell(3, 10).Range.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    }
                    else if ((radioButton2.Checked))
                    {
                        button2.Visible = false;
                        int y;
                        progressBar1.Visible = true;
                        progressBar1.Minimum = 0;
                        progressBar1.Maximum = main.dataGridView1.RowCount;
                        progressBar1.Step = 1;
                        for (int i = 0; i < main.dataGridView1.RowCount; i++)
                        {
                            for (int j = 1; j < 11; j++)
                            {
                                try
                                {
                                    if (j == 5)
                                    {
                                        DateTime birthday = (DateTime)main.dataGridView1[j, i].Value;
                                        oTable.Cell(i + 3, j).Range.Text = birthday.ToShortDateString();
                                    }
                                    else
                                        oTable.Cell(i + 3, j).Range.Text = main.dataGridView1[j, i].Value.ToString();
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show("Невозможно сформировать файл. " + ex.Message);
                                }
                                if (j == 1)
                                {
                                    oTable.Cell(i + 3, j).Range.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                                }
                                if (j == 10)
                                {
                                    oTable.Cell(i + 3, j).Range.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                                }
                            }
                            progressBar1.PerformStep();
                            y = 100 * i / main.dataGridView1.RowCount + 1;
                            label1.Text = y.ToString() + " %";
                        }
                        progressBar1.Value = 0;
                        for (int i = 3; i < main.dataGridView1.RowCount + 3; i++)
                        {
                            oTable.Cell(i, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                            oTable.Cell(i, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                            oTable.Cell(i, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                            oTable.Cell(i, 6).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                            oTable.Cell(i, 7).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                            oTable.Cell(i, 10).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                            progressBar1.PerformStep();
                            y = 100 * i / main.dataGridView1.RowCount;
                            label1.Text = y.ToString() + " %";
                        }
                     }
                    oTable.Range.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    oTable.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent);

                    Word.Window activeWindow = oDoc.Application.ActiveWindow;
                    object currentPage = Word.WdFieldType.wdFieldPage;
                    activeWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageFooter;
                    activeWindow.ActivePane.Selection.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    activeWindow.Selection.Fields.Add(activeWindow.Selection.Range, ref currentPage, ref oMissing, ref oMissing);
                    activeWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekMainDocument;

                    oDoc.SaveAs(fileName, ref oMissing, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing);
                    oDoc.Close(ref oMissing, ref oMissing, ref oMissing);
                    oWord.Quit(ref oMissing, ref oMissing, ref oMissing);  
                 }
                else if ((radioButton5.Checked))
                {
                    if ((radioButton1.Checked))
                    {
                        File.WriteAllText(@fileName,  GlobalData.Tab + " " +
                                                      GlobalData.F + " " +
                                                      GlobalData.I + " " +
                                                      GlobalData.O + " " +
                                                      GlobalData.Bithday.ToShortDateString() + " " +
                                                      GlobalData.Otdel + " " +
                                                      GlobalData.Dolzhn + " " +
                                                      GlobalData.Tel + " " +
                                                      GlobalData.Comp + " " +
                                                      GlobalData.Email, Encoding.GetEncoding(1251));
                    }
                    else if ((radioButton2.Checked))
                    {
                        button2.Visible = false;
                        int y;
                        progressBar1.Visible = true;
                        progressBar1.Minimum = 0;
                        progressBar1.Maximum = main.dataGridView1.RowCount;
                        progressBar1.Step = 1;
                        FileInfo f = new FileInfo(@fileName);
                        string stroka = "";
                        using (FileStream fs = f.Open(FileMode.OpenOrCreate, FileAccess.Write, FileShare.Read)) 
                        { 
                        using (StreamWriter sw = new StreamWriter(fs, Encoding.GetEncoding(1251))) 
                            {
                                for (int i = 0; i < main.dataGridView1.RowCount; i++)
                                {
                                    for (int j = 1; j < 11; j++)
                                    {
                                        if (j == 5)
                                        {
                                            DateTime birthday = (DateTime)main.dataGridView1[j, i].Value;
                                            stroka += birthday.ToShortDateString() + " ";
                                        }
                                        else
                                            stroka += main.dataGridView1[j, i].Value.ToString() + " ";
                                    }
                                    sw.WriteLine(stroka);
                                    stroka = "";
                                    progressBar1.PerformStep();
                                    y = 100 * i / main.dataGridView1.RowCount + 1;
                                    label1.Text = y.ToString() + " %";
                                }
                            } 
                        } 
                    }
                }
                MessageBox.Show("Файл сформирован и записан", " ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk );
                this.Close();
            }
            else
            {
                return;
            }
        }


//==============================================================================================
// Очистка памяти после закрытия Excel
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.GetTotalMemory(true);
            }
        }
        

//==============================================================================================
// Делает активным последующий выбор в зависимости от предыдущего
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
                groupBox2.Enabled = true;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
                groupBox2.Enabled = true;
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton3.Checked)
            {
                groupBox3.Enabled = true;
                filter = "Книга Excel (*.xls)|*.xls";
            }
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton4.Checked)
            {
                groupBox3.Enabled = true;
                filter = "Документ Word (*.doc)|*.doc";
            }
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton5.Checked)
            {
                groupBox3.Enabled = true;
                filter = "Обычный текст (*.txt)|*.txt";
            }
        }

      
    }
}
