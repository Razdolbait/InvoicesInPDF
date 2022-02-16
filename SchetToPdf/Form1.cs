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
        public class invoice
        {
            public int Id;
            public int UNP;
            public List<sheet> Text;
        }
        public class sheet
        {
            public int Orient;// Ориентация страницы 0 - книжная, 1 - альбомная
            public List<string> Text;
        }
        enum strInPage //кол-во строк на странице
        {
            Landscape = 49,
            Portrait = 68
        };

        List<invoice> invoices = new List<invoice>();

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Текстовые файлы(*.txt) | *.txt";
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            string directoryRez = Path.GetDirectoryName(openFileDialog1.FileName);
            string[] allfiles = Directory.GetFiles(directoryRez, "*.txt");
            foreach (var f in allfiles)
            {
                string[] fileText = File.ReadAllLines(f, Encoding.GetEncoding(866));//считываем весь файл построчно
                int orient;// Ориентация страницы 0 - книжная, 1 - альбомная
                if (Path.GetFileNameWithoutExtension(f).Substring(Path.GetFileNameWithoutExtension(f).Length - 1) == "_")
                {
                    orient = 1;
                }
                else
                {
                    orient = 0;
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
                i = 0;
                //формируем счета
                foreach (var s in arrListString)
                {
                    var id = s[0].Trim().Substring(1, 9);
                    s.Remove(s[0]);
                    List<string> tempList = new List<string>();
                    int arrSize;
                    if (orient == 0)
                    {
                        arrSize = (int)strInPage.Portrait;
                    }
                    else
                    {
                        arrSize = (int)strInPage.Landscape;
                    }
                    //разбивка на страницы
                    while (s.Count > 0)
                    {
                        string temp = "";
                        if (s.Count > arrSize)
                        {
                            for (int n = 0; n < arrSize; n++)
                            {
                                temp += s[0] + '\r' + '\n';
                                s.RemoveAt(0);
                            }
                        }
                        else
                        {
                            while (s.Count > 0)
                            {
                                temp += s[0] + '\r' + '\n';
                                s.RemoveAt(0);
                            }
                        }
                        tempList.Add(temp);
                    }
                    //группировка счетов по УНП
                    if (int.TryParse(id, out int number))
                    {
                        if(invoices.Exists(x => x.UNP == number))
                        {
                            invoices[invoices.FindIndex(x => x.UNP == number)].Text.Add(
                                new sheet
                                {
                                    Orient = orient,
                                    Text = tempList
                                });
                            
                        }
                        else
                        {
                            invoices.Add(new invoice
                            {
                                Id = i,
                                UNP = number,
                                Text = new List<sheet>
                                {
                                    new sheet
                                    {
                                        Orient = orient,
                                        Text = tempList
                                    }
                                }
                            });
                        }
                        i++;
                    }
                }

            }
            //создаем pdf'ки
            foreach(var s in invoices)
            {
                PdfDocument document = new PdfDocument();
                XRect rect;
                List<string> textList = new List<string>();
                foreach (var ss in s.Text)
                {
                    if (ss.Orient == 0)
                    {
                        rect = new XRect(20, 20, 570, 800);
                    }
                    else
                    {
                        rect = new XRect(20, 20, 800, 570);
                    }
                    foreach(var t in ss.Text)
                    {
                        PdfPage page = document.AddPage();
                        if (ss.Orient == 0)
                        {
                            page.Orientation = PdfSharp.PageOrientation.Portrait;
                        }
                        else
                        {
                            page.Orientation = PdfSharp.PageOrientation.Landscape;
                        }
                        XGraphics gfx = XGraphics.FromPdfPage(page);
                        XFont font = new XFont("Courier New", 10, XFontStyle.Regular);
                        XTextFormatter tf = new XTextFormatter(gfx);
                        tf.Alignment = XParagraphAlignment.Left;
                        tf.DrawString(t, font, XBrushes.Black, rect, XStringFormat.TopLeft);
                    }
                }
                document.Save(directoryRez + @"\" + s.UNP.ToString() + ".pdf");
            }
            MessageBox.Show("Готово!", "Уведомление", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
