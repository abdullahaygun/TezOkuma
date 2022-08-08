using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Core;
using word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Diagnostics;
using System.Threading;

namespace WordOkuma
{
    public partial class Form1 : Form
    {
        OpenFileDialog file = new OpenFileDialog();
        word.Application app;
        word.Document doc;
        word.Range range;
        word.Find find;
        object dosya;
        object nullobject;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            dosyaSec();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            kapat();
        }

        public void dosyaSec()
        {

            file.Filter = "Word Belgesi(.docx) | *.docx| Word Belgesi(.doc) | *.doc";
            if (file.ShowDialog() == DialogResult.OK)
            {
                dosya = file.FileName;
                app = new word.Application();
                nullobject = System.Reflection.Missing.Value;
                doc = app.Documents.Open(dosya, nullobject, false, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, nullobject, false, nullobject, nullobject, nullobject, nullobject);
            }
        }

        public void temizle(params RichTextBox[] rich)
        {
            for (int i = 0; i < rich.Length; i++)
            {
                rich[i].Clear();
            }
        }

        public void kapat()
        {
            if (app != null)
            {
                //app.ActiveDocument.Close();
                doc.Close(nullobject, nullobject, nullobject);
                app.Quit(nullobject, nullobject, nullobject);

            }
        }

        public string aralikOkuText(int start, int end)
        {
            string aralik = "";
            aralik = doc.Range(start, end).Text.ToString();
            return aralik;
        }

        public string paragrafOku(int basParagraf, int sonParagraf)
        {
            string satir = "";
            if (doc != null)
            {
                int i = basParagraf;

                foreach (word.Paragraph objParagraph in doc.Paragraphs)
                {

                    if (i == sonParagraf)
                    {
                        break;
                    }
                    else
                    {
                        satir += doc.Paragraphs[i + 1].Range.Text + "\n";
                        i++;
                    }

                }
            }

            return satir;

        }

        public string paragrafOku()
        {
            string satir = "";
            if (doc != null)
            {
                int i = 0;
                foreach (word.Paragraph objParagraph in doc.Paragraphs)
                {
                    satir += doc.Paragraphs[i + 1].Range.Text + "\n";
                    i++;
                }
            }
            return satir;
        }

        private void btnPageMarginControl_Click(object sender, EventArgs e)
        {
            richTextBox1.Text = sayfaKenarBosluk();
        }

        public void tumBelgeyiOku(RichTextBox rich)
        {
            if (doc != null)
            {
                doc.ActiveWindow.Selection.WholeStory();
                doc.ActiveWindow.Selection.Copy();
                IDataObject data = Clipboard.GetDataObject();
                rich.Rtf = data.GetData(DataFormats.Rtf).ToString();

            }

        }

        public void istatistikVeriler(RichTextBox rich)
        {
            if (doc != null)
            {
                word.WdStatistic pageCountStat = word.WdStatistic.wdStatisticPages;
                word.WdStatistic statTotalChar = word.WdStatistic.wdStatisticCharacters;
                word.WdStatistic statTotoalLines = word.WdStatistic.wdStatisticLines;
                word.WdStatistic statTotalParagraf = word.WdStatistic.wdStatisticParagraphs;
                word.WdStatistic statTotalWords = word.WdStatistic.wdStatisticWords;
                word.WdStatistic statTotalCharBosluk = word.WdStatistic.wdStatisticCharactersWithSpaces;


                int karakter = doc.ComputeStatistics(statTotalChar, false);
                int satirr = doc.ComputeStatistics(statTotoalLines, false);
                int paragraf = doc.ComputeStatistics(statTotalParagraf, false);
                int kelime = doc.ComputeStatistics(statTotalWords, false);
                int karakterBoslukla = doc.ComputeStatistics(statTotalCharBosluk, false);
                int pageCount = doc.ComputeStatistics(pageCountStat, ref nullobject);

                int tablo = 0;
                foreach (word.Table objTable in doc.Tables)
                {
                    tablo++;

                }

                int resim = 0;
                foreach (word.InlineShape objResim in doc.InlineShapes)
                {
                    resim++;

                }

                string digerIcindekiler = doc.TablesOfFigures.Count.ToString();

                rich.Text += "Diğer İçindekiler Sayısı: " + digerIcindekiler.ToString() + " \n";
                rich.Text += "Toplam RESİM sayısı: " + resim.ToString() + "\n";
                rich.Text += "Toplam TABLO sayısı: " + tablo.ToString() + "\n";
                rich.Text += "Toplam Karakter: " + karakter.ToString() + " \n";
                rich.Text += "Toplam Satır: " + satirr.ToString() + " \n";
                rich.Text += "Toplam Paragraf: " + paragraf.ToString() + " \n";
                rich.Text += "Toplam Kelime: " + kelime.ToString() + " \n";
                rich.Text += "Toplam Karakter(Boşluk Dahil): " + karakterBoslukla.ToString() + " \n";
                rich.Text += "Toplam SAYFA: " + pageCount.ToString() + " \n";
            }

        }

        private void btnOku_Click(object sender, EventArgs e)
        {
            temizle(richTextBox1, richTextBox2);
            tumBelgeyiOku(richTextBox1);
            istatistikVeriler(richTextBox2);
        }

        public string sayfaKenarBosluk()
        {
            string cikti = "";
            if (app != null)
            {
                float topMargin = doc.PageSetup.TopMargin;
                float bottomMargin = doc.PageSetup.BottomMargin;
                float leftMargin = doc.PageSetup.LeftMargin;
                float rightMargin = doc.PageSetup.RightMargin;
                string orient = doc.PageSetup.Orientation.ToString();
                float ustKenarBoslugu = -1;
                float altKenarBoslugu = -1;
                float solKenarBoslugu = -1;
                float sagKenarBoslugu = -1;

                if (orient == "wdOrientPortrait")
                {
                    ustKenarBoslugu = app.CentimetersToPoints(3f);
                    altKenarBoslugu = app.CentimetersToPoints(2.5f);
                    solKenarBoslugu = app.CentimetersToPoints(3.25f);
                    sagKenarBoslugu = app.CentimetersToPoints(2.5f);


                    if (Math.Round(app.PointsToCentimeters(topMargin), 2) != app.PointsToCentimeters(ustKenarBoslugu))
                    {
                        cikti += "ÜST Kenar Boşluğunuz: " + Math.Round(app.PointsToCentimeters(topMargin), 2).ToString() + " ANCAK " + app.PointsToCentimeters(ustKenarBoslugu).ToString() + " olmalıdır!" + "\n";
                    }

                    if (Math.Round(app.PointsToCentimeters(bottomMargin), 2) != app.PointsToCentimeters(altKenarBoslugu))
                    {
                        cikti += "ALT Kenar Boşluğunuz: " + Math.Round(app.PointsToCentimeters(bottomMargin), 2).ToString() + " ANCAK " + app.PointsToCentimeters(altKenarBoslugu).ToString() + " olmalıdır!" + "\n";
                    }

                    if (Math.Round(app.PointsToCentimeters(leftMargin), 2) != app.PointsToCentimeters(solKenarBoslugu))
                    {
                        cikti += "SOL Kenar Boşluğunuz: " + Math.Round(app.PointsToCentimeters(leftMargin), 2).ToString() + " ANCAK " + app.PointsToCentimeters(solKenarBoslugu).ToString() + " olmalıdır!" + "\n";
                    }

                    if (Math.Round(app.PointsToCentimeters(rightMargin), 2) != app.PointsToCentimeters(sagKenarBoslugu))
                    {
                        cikti += "SAĞ Kenar Boşluğunuz: " + Math.Round(app.PointsToCentimeters(rightMargin), 2).ToString() + " ANCAK " + app.PointsToCentimeters(sagKenarBoslugu).ToString() + " olmalıdır!" + "\n";

                    }
                }
                else if (orient == "wdOrientLandscape")
                {
                    ustKenarBoslugu = app.CentimetersToPoints(3.25f);
                    altKenarBoslugu = app.CentimetersToPoints(2.5f);
                    solKenarBoslugu = app.CentimetersToPoints(2.5f);
                    sagKenarBoslugu = app.CentimetersToPoints(3.0f);


                    if (Math.Round(app.PointsToCentimeters(topMargin), 2) != app.PointsToCentimeters(ustKenarBoslugu))
                    {
                        cikti += "ÜST Kenar Boşluğunuz: " + Math.Round(app.PointsToCentimeters(topMargin), 2).ToString() + " ANCAK " + app.PointsToCentimeters(ustKenarBoslugu).ToString() + " olmalıdır!" + "\n";
                    }

                    if (Math.Round(app.PointsToCentimeters(bottomMargin), 2) != app.PointsToCentimeters(altKenarBoslugu))
                    {
                        cikti += "ALT Kenar Boşluğunuz: " + Math.Round(app.PointsToCentimeters(bottomMargin), 2).ToString() + " ANCAK " + app.PointsToCentimeters(altKenarBoslugu).ToString() + " olmalıdır!" + "\n";
                    }

                    if (Math.Round(app.PointsToCentimeters(leftMargin), 2) != app.PointsToCentimeters(solKenarBoslugu))
                    {
                        cikti += "SOL Kenar Boşluğunuz: " + Math.Round(app.PointsToCentimeters(leftMargin), 2).ToString() + " ANCAK " + app.PointsToCentimeters(solKenarBoslugu).ToString() + " olmalıdır!" + "\n";
                    }

                    if (Math.Round(app.PointsToCentimeters(rightMargin), 2) != app.PointsToCentimeters(sagKenarBoslugu))
                    {
                        cikti += "SAĞ Kenar Boşluğunuz: " + Math.Round(app.PointsToCentimeters(rightMargin), 2).ToString() + " ANCAK " + app.PointsToCentimeters(sagKenarBoslugu).ToString() + " olmalıdır!" + "\n";
                    }
                }
            }
            return cikti;
        }

        public void Basliklar()
        {
            object baslik = "";
            for (int i = 1; i < app.CaptionLabels.Count; i++)
            {

                baslik += app.CaptionLabels[i].Name + "\n";
            }

            for (int i = 0; i < doc.Bibliography.Sources.Count; i++)
            {

                baslik += doc.Bibliography.Sources[i].ToString();
            }


            richTextBox1.Text = baslik.ToString();
        }

        private void btnBaslikKontrol_Click(object sender, EventArgs e)
        {

            TemelYaziDuzeni();
        }

        public void kelimeAra(string aranacakİfade)
        {
            range = doc.Content;
            find = range.Find;
            find.Text = aranacakİfade;

            if (find.Execute())
            {
                MessageBox.Show("Bulundu");
            }
        }

        public void ParagrafYapisi()
        {
            string satir = "";
            if (doc != null)
            {

                foreach (word.Paragraph objParagraph in doc.Paragraphs)
                {


                }
            }

            richTextBox1.Text += satir;
        }

        public void TemelYaziDuzeni()
        {

            int count = doc.Words.Count;
            pBar.Maximum = count;
            //object[] text = new object[count];
            //object[] font = new object[count];
            //object[] size = new object[count];
            //object[] bold = new object[count];
            //object[] italic = new object[count];
            //object[] color = new object[count];
            List<object> line = new List<object>(count);
            List<object> text = new List<object>(count);
            List<object> font = new List<object>(count);
            List<object> size = new List<object>(count);
            List<object> bold = new List<object>(count);
            List<object> italic = new List<object>(count);
            List<object> color = new List<object>(count);
            List<object> hiza = new List<object>(count);
            List<object> lineSpace = new List<object>(count);
            List<object> shapes = new List<object>(count);
            List<object> pages = new List<object>(count);
            //for (int i = 1; i < doc.Paragraphs.Count; i++) // Paragraf sayıp hizaları, satır aralıkları kaydediyor.
            //{
            //    object tempLine = doc.Paragraphs[i].Range.Text;
            //    object tempAlignment = doc.Paragraphs[i].Alignment;
            //    object tempLineSpacing = doc.Paragraphs[i].LineSpacingRule;
            //    object tempPages = word.WdFieldType.wdFieldPage;
            //    if ((doc.Paragraphs[i].Range.End - doc.Paragraphs[i].Range.Start) > 1)
            //    {
            //        line.Add(tempLine.ToString());
            //        hiza.Add(tempAlignment.ToString());
            //        lineSpace.Add(tempLineSpacing.ToString());
            //        pages.Add(tempPages.ToString());
            //    }
            //}
            var timer = new Stopwatch();
            timer.Start();
            //Task.Factory.StartNew(() =>
            new Thread(() =>
            {
                for (int i = 1; i < count; i++)
                {

                    //System.GC.Collect();
                    //System.GC.WaitForPendingFinalizers();

                    //object tempText = "";
                    //object tempFont = "";
                    //object tempSize = "";
                    //object tempBold = "";
                    //object tempItalic = "";
                    //object tempColor = "";

                    //if (doc.Words[i].Text != null)
                    //{
                    //    tempText = doc.Words[i].Text;
                    //    tempFont = doc.Words[i].Font.Name;
                    //    tempSize = doc.Words[i].Font.Size;
                    //    tempBold = doc.Words[i].Font.Bold;
                    //    tempItalic = doc.Words[i].Font.Italic;
                    //    tempColor = doc.Words[i].Font.Color;
                    //}

                    //if (doc.Words[i].Text != null)
                    //{
                    //text[i] = tempText;
                    //font[i] = tempFont;
                    //size[i] = tempSize;
                    //bold[i] = tempBold;
                    //italic[i] = tempItalic;
                    //color[i] = tempColor;
                    text.Add(doc.Words[i].Text);
                    font.Add(doc.Words[i].Font.Name);
                    size.Add(doc.Words[i].Font.Size);
                    bold.Add(doc.Words[i].Font.Bold);
                    italic.Add(doc.Words[i].Font.Italic);
                    color.Add(doc.Words[i].Font.Color);
                    //}


                    //pBar.PerformStep();
                }

                for (int i = 1; i < 5; i++)
                {
                    richTextBox1.Text += text[i] + " : " + font[i] + " : " + size[i] + " : " + bold[i] + " : " + italic[i] + " : " + color[i] + "\n";
                }
                //});
            }).Start();
            this.Text = timer.Elapsed.TotalSeconds.ToString();

            //for (int i = 1; i < doc.Shapes.Count; i++)
            //{
            //    object tempShapes = doc.Shapes[i].AnchorID;
            //    shapes.Add(tempShapes.ToString());
            //}


            //for (int i = 1; i < count; i++) // Kelime kelime tarıyor.
            //{
            //    if (doc.Words[i].Text != null)
            //    {
            //        richTextBox1.Text += text[i] + " : " + font[i] + " : " + size[i] + " : " + bold[i] + " : " + italic[i] + " : " + color[i] + "\n";
            //    }





            //    //if (text[i].ToString() + " " + text[i + 1].ToString() + text[i + 2].ToString() + text[i + 3].ToString() + text[i + 4].ToString() == "Şekil 2.4.")
            //    //{
            //    //    MessageBox.Show(i + ". kelimede bulundu");
            //    //}
            //}
            //istatistikVeriler(richTextBox2);

            //for (int i = 0; i < line.Count; i++) // Paragraf paragraf tarıyor.
            //{
            //    richTextBox1.Text += line[i] + " : " + hiza[i] + " : " + lineSpace[i] + "\n";
            //    if (line[i].ToString().Contains("Şekil 2.4.") == true)
            //    {
            //        MessageBox.Show((i + 1) + ". satırda bulundu");
            //    }
            //}


            //for (int i = 0; i < shapes.Count; i++) // Resim resim tarıyor.
            //{
            //    richTextBox1.Text += shapes[i] + "\n";
            //}

            //sqlLite sql = new sqlLite();
            //sql.olusturDataBase();
            //sql.veriEkle(doc, 200, pBar);

        }
    }
}
