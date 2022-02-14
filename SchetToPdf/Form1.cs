using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Orientation = MigraDoc.DocumentObjectModel.Orientation;

namespace SchetToPdf
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        public class receipt
        {
            public int Id;
            public string UNP;
            public List<string> Text;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Текстовые файлы(*.txt) | *.txt";
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            string fileNameRez = Path.GetFileNameWithoutExtension(openFileDialog1.FileName);
            string directoryRez = Path.GetDirectoryName(openFileDialog1.FileName);
            string[] fileText = File.ReadAllLines(openFileDialog1.FileName, Encoding.GetEncoding(866));//считываем весь файл построчно

            int orient;// Ориентация страницы 0 - книжная, 1 - альбомная
            int strInPage; //кол-во строк на странице
            if (fileNameRez.Substring(fileNameRez.Length - 1) == "_")
            {
                orient = 1;
                strInPage = 49;
            }
            else
            {
                orient = 0;
                strInPage = 68;
            }

            //разбираем файл на блоки-счета
            List<string>[] arrListString = new List<string>[0];
            int i = -1;
            foreach (var s in fileText)
            {
                if (s.Trim().Length > 0)
                {
                    if (s.Trim().Substring(0, 1) == "<")
                    {
                        i++;
                        Array.Resize(ref arrListString, i + 1);
                        arrListString[i] = new List<string>();
                        arrListString[i].Add(s);
                    }
                    else arrListString[i].Add(s);
                }
                else arrListString[i].Add(s);
            }
            //формируем счета
            i = 0;
            List<receipt> receipts = new List<receipt>();
            foreach(var s in arrListString)
            {
                var id = s[0].Trim().Substring(1, 9);
                s[0] = "";
                receipts.Add(new receipt
                {
                    Id = i,
                    UNP = id,
                    Text = s
                });
                i++;
            }
            //создаем pdf'ки
            //foreach(var s in receipts)
            //{
            //    Document document = new Document();
            //    Section section = document.AddSection();
            //    section.PageSetup.PageFormat = PageFormat.A4;//стандартный размер страницы
            //    //Ориентация страницы зависит от последнего символа в имени файла
            //    //name - книжная 
            //    //name_ - альбомная
            //    if (fileNameRez.Substring(fileNameRez.Length - 1) == "_")
            //    {
            //        section.PageSetup.Orientation = Orientation.Landscape;
            //    }
            //    else section.PageSetup.Orientation = Orientation.Portrait;
            //    section.PageSetup.BottomMargin = 20;//нижний отступ
            //    section.PageSetup.TopMargin = 20;//верхний отступ
            //    section.PageSetup.LeftMargin = 25;
            //    section.PageSetup.RightMargin = 15;
            //    foreach (var t in s.Text)
            //    {
            //        Paragraph paragraph = new Paragraph();
            //        paragraph.Format.Font.Name = "Courier New";
            //        paragraph.Format.Font.Size = 10;
            //        section.Add(paragraph);
            //        paragraph.AddFormattedText(t/*.Replace(' ', Convert.ToChar("\u202F"))*/);
            //    }
            //    PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(true, PdfFontEmbedding.Always);
            //    pdfRenderer.Document = document;
            //    pdfRenderer.RenderDocument();
            //    pdfRenderer.PdfDocument.Save(directoryRez + @"\" + s.UNP.ToString() + '_' + fileNameRez + ".pdf");// сохраняем
            //}
           foreach(var s in receipts)
            {
                PdfDocument document = new PdfDocument();
                XRect rect;
                if (orient == 0)
                {
                    rect = new XRect(20, 20, 570, 800);
                }
                else rect = new XRect(20, 20, 800, 570);
                List<string> textList = new List<string>();
                i = 0;
                string tempStr = "";
                foreach(var t in s.Text)
                {
                    i++;
                    if(i%strInPage == 0)
                    {
                        textList.Add(tempStr);
                        tempStr = t + '\r' + '\n';
                        continue;
                    }
                    tempStr += t + '\r' + '\n';
                }
                textList.Add(tempStr);
                foreach (var tl in textList)
                {
                    PdfPage page = document.AddPage();
                    if (orient == 0)
                    {
                        page.Orientation = PdfSharp.PageOrientation.Portrait;
                    }
                    else page.Orientation = PdfSharp.PageOrientation.Landscape;
                    XGraphics gfx = XGraphics.FromPdfPage(page);
                    XFont font = new XFont("Courier New", 10, XFontStyle.Regular);
                    XTextFormatter tf = new XTextFormatter(gfx);
                    tf.Alignment = XParagraphAlignment.Left;
                    tf.DrawString(tl, font, XBrushes.Black, rect, XStringFormat.TopLeft);
                }
                document.Save(directoryRez + @"\" + s.UNP.ToString() + '_' + fileNameRez + ".pdf");
            }

            //PdfDocument document = new PdfDocument();
            //PdfPage page = document.AddPage();

            //else page.Orientation = PdfSharp.PageOrientation.Portrait;
            //XGraphics gfx = XGraphics.FromPdfPage(page);
            //XFont font = new XFont("Courier New", 10, XFontStyle.Regular);//Coueer New
            //XTextFormatter tf = new XTextFormatter(gfx);
            //tf.Alignment = XParagraphAlignment.Left;
            //XRect rect = new XRect(20, 20, 1000, page.Height - 20);
            //tf.DrawString(s.Text, font, XBrushes.Black, rect, XStringFormat.TopLeft);
            //document.Save(directoryRez + @"\" + s.UNP.ToString() + '_' + fileNameRez + ".pdf");
            MessageBox.Show("Готово!", "Уведомление", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
