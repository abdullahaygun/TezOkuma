using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace Ribbon
{
    class Kontroller
    {

        public void BaslikVarMi(string baslikAdi, string fontName, int punto, int bold, int buyukKucuk, int sonraBosluk, WdLineSpacing satirAralık, WdParagraphAlignment hiza)
        {
            // Aktif olan doküman ile bağlantı sağlayabilmek için "wordApp" Application nesnesi oluşturuldu.
            var wordApp = (_Application)Marshal.GetActiveObject("Word.Application");

            // Tez kurallarına uygun arama paramtreleri tanımlandı.
            wordApp.Selection.Find.Font.Name = fontName;
            wordApp.Selection.Find.Font.Size = punto;
            wordApp.Selection.Find.Font.Bold = bold;
            wordApp.Selection.Find.Font.SmallCaps = buyukKucuk;
            wordApp.Selection.Find.ParagraphFormat.Alignment = hiza;
            wordApp.Selection.Find.ParagraphFormat.SpaceAfter = sonraBosluk;
            wordApp.Selection.Find.ParagraphFormat.SpaceBefore = 0;
            wordApp.Selection.Find.ParagraphFormat.LineSpacingRule = satirAralık;
            wordApp.Selection.Find.Execute(baslikAdi);

            // Belge içerisinde tanımlanmış olan uygun parametrelere göre arama yapılıp bulunan kelimelerin sayısını veren Integer tipinde "foundNumber" değişkeni tanımlandı.
            int foundNumber = 0;

            // Verilen parametrelere uygun olarak arama yapıldığında bulunan Başlık sayısını veren döngü oluşturuldu.
            while (wordApp.Selection.Find.Found)
            {
                foundNumber++;
                wordApp.Selection.Find.Execute(baslikAdi);
            }

            // "Microsoft.Office.Interop.Word" kütüphanesinin "Selection" methodunu kullanarak bulunan kelimeleri seçtikten sonra cursor 'u "RESETLEMEK" için 1. sayfanın 1.satırına getirlmesi sağlandı.
            wordApp.Selection.GoTo(Microsoft.Office.Interop.Word.WdGoToItem.wdGoToPage, Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToAbsolute, 1).Select();
            wordApp.Selection.GoTo(Microsoft.Office.Interop.Word.WdGoToItem.wdGoToLine, Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToAbsolute, 1).Select();

            // Kullanıcı, alacağı çıktının anlamlı olabilmesi adına doğru olarak yazdığı Başlığın Sayfa ve Satır Numarası getirilmesi için değişkenler(startLine, startPageNum) tanımlandı.
            int startLine = 0;
            int startPageNum = 0;

            // Tanımlanmış olan "startLine" ve "startPageNum" değişkenlerine değerler atandı.
            for (int i = 0; i < foundNumber; i++)
            {
                startLine = wordApp.Selection.Information[WdInformation.wdFirstCharacterLineNumber];
                startPageNum = wordApp.Selection.Information[WdInformation.wdActiveEndPageNumber];

            }

            // Eğer Belge üzerinde yazılan Başlık ve başlığa ait özellikten bir adet varsa ya da yoksa uygun bir çıktı verilmesi sağlandı.
            if (foundNumber == 1)
            {
                MessageBox.Show(baslikAdi + " Başlığı " + startPageNum + ". sayfada " + startLine + ". satırda bulundu..");
            }
            else
            {
                MessageBox.Show(baslikAdi + " Başlığı verilen kurallara göre BULUNAMADI!");
            }

            // "Microsoft.Office.Interop.Word" kütüphanesinin "Selection" methodunu kullanarak bulunan kelimeleri seçtikten sonra cursor 'u "RESETLEMEK" için 1. sayfanın 1.satırına getirlmesi sağlandı.
            wordApp.ActiveDocument.GoTo(Microsoft.Office.Interop.Word.WdGoToItem.wdGoToPage, Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToAbsolute, 1).Select();
            wordApp.ActiveDocument.GoTo(Microsoft.Office.Interop.Word.WdGoToItem.wdGoToLine, Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToAbsolute, 1).Select();
        }


        public void onBolumBaslikkKontrol(string fontName, int punto, int bold, int buyukKucuk, int onceBosluk, WdLineSpacing satirAralık, WdParagraphAlignment hiza)
        {
            var wordApp = (_Application)Marshal.GetActiveObject("Word.Application");
            Microsoft.Office.Interop.Word.Range range = wordApp.ActiveDocument.Content;
            //Regex rKaynakBulucu = new Regex(@"\[(.\d)\]");
            Regex rKaynakBulucu = new Regex(@"\[\d*?\]");
            List<string> kaynakList = new List<string>();
            //List<string> icerikKaynakList = new List<string>();
            List<object> icerikKaynakList = new List<object>();

            range.Find.Font.Name = fontName;
            range.Find.Font.Size = punto;
            //range.Find.Font.Bold = bold;
            //range.Find.Font.SmallCaps = buyukKucuk;
            //range.Find.ParagraphFormat.Alignment = hiza;
            //range.Find.ParagraphFormat.SpaceBefore = onceBosluk;
            //range.Find.ParagraphFormat.LineSpacingRule = satirAralık;


            var rangeBas = -1;
            var rangeSon = -1;
            var icerik = wordApp.ActiveDocument.Range(0, 0);
            string baslik = "";
            int k = 0;
            int sayac = 0;
            while (range.Find.Execute(""))
            {
                if (range.Find.Found)
                {
                    if ((range.End - range.Start) > 1)
                    {
                        sayac++;
                        if (sayac == 1)
                        {
                            baslik = range.Text;
                            rangeBas = range.Start + range.Text.Length;
                            int geriDonus = 0;
                            while (range.Find.Execute(""))
                            {
                                geriDonus++;

                                if ((range.End - range.Start) > 1)
                                {

                                    rangeSon = range.End - range.Text.Length;
                                    break;
                                }
                            }
                            icerik = wordApp.ActiveDocument.Range(rangeBas, rangeSon);

                            range.Find.Forward = false;
                            for (int i = 0; i < geriDonus; i++)
                            {
                                range.Find.Execute("");

                            }
                            range.Find.Forward = true;
                            //MessageBox.Show(icerik.Text+"asdasd");
                        }
                        else
                        {

                            baslik = range.Text;

                            rangeBas = range.Start + range.Text.Length;
                            int geriDonus2 = 0;
                            while (range.Find.Execute(""))
                            {
                                geriDonus2++;

                                if ((range.End - range.Start) > 1)
                                {

                                    rangeSon = range.End - range.Text.Length;
                                    break;
                                }
                            }
                            icerik = wordApp.ActiveDocument.Range(rangeBas, rangeSon);

                            range.Find.Forward = false;
                            for (int i = 0; i < geriDonus2; i++)
                            {
                                range.Find.Execute("");
                                // MessageBox.Show(geriDonus2.ToString() + " asd: " + range.Text);
                            }
                            range.Find.Forward = true;
                        }


                        if (rangeBas < rangeSon)
                        {
                            icerik.Select();
                            icerik.Copy();
                            IDataObject data = Clipboard.GetDataObject();
                            string hafıza = data.GetData(DataFormats.Text).ToString();

                            if (!baslik.Trim().Equals("KAYNAKLAR"))
                            {
                                icerikKaynakList.Add(rKaynakBulucu.Match(hafıza));
                                MessageBox.Show(icerikKaynakList[0].ToString());




                            }
                            else
                            {
                                for (int i = 0; i < rKaynakBulucu.Matches(hafıza).Count; i++)
                                {
                                    kaynakList.Add(rKaynakBulucu.Matches(hafıza)[i].Value);
                                    //MessageBox.Show(kaynakList[i]);
                                }


                            }



                        }

                        k++;
                    }

                }





            }

            for (int i = 0; i < kaynakList.Count; i++)
            {
                MessageBox.Show(kaynakList[i]);
            }

        }
        public void sayfaKenarBoslukveA4Kontrol()
        {
            var wordApp = (_Application)Marshal.GetActiveObject("Word.Application");
            string cikti = "";
            Microsoft.Office.Interop.Word.Range range = wordApp.ActiveDocument.Content;
            Microsoft.Office.Interop.Word.WdStatistic pageCountStat = Microsoft.Office.Interop.Word.WdStatistic.wdStatisticPages;
            int pageCount = wordApp.ActiveDocument.ComputeStatistics(pageCountStat);
            for (int i = 1; i <= pageCount; i++)
            {
                wordApp.Selection.GoTo(Microsoft.Office.Interop.Word.WdGoToItem.wdGoToPage, Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToAbsolute, i).Select();
                // wordApp.Selection.GoTo(Microsoft.Office.Interop.Word.WdGoToItem.wdGoToLine, Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToAbsolute, 1).Select();



                if (wordApp != null)
                {
                    float height = wordApp.Selection.PageSetup.PageHeight;
                    float width = wordApp.Selection.PageSetup.PageWidth;
                    float topMargin = wordApp.Selection.PageSetup.TopMargin;
                    float bottomMargin = wordApp.Selection.PageSetup.BottomMargin;
                    float leftMargin = wordApp.Selection.PageSetup.LeftMargin;
                    float rightMargin = wordApp.Selection.PageSetup.RightMargin;
                    string orient = wordApp.Selection.PageSetup.Orientation.ToString();
                    float ustKenarBoslugu = -1;
                    float altKenarBoslugu = -1;
                    float solKenarBoslugu = -1;
                    float sagKenarBoslugu = -1;
                    float yuk = -1;
                    float gen = -1;


                    if (orient == "wdOrientPortrait")
                    {
                        ustKenarBoslugu = wordApp.CentimetersToPoints(3f);
                        altKenarBoslugu = wordApp.CentimetersToPoints(2.5f);
                        solKenarBoslugu = wordApp.CentimetersToPoints(3.25f);
                        sagKenarBoslugu = wordApp.CentimetersToPoints(2.5f);
                        yuk = wordApp.CentimetersToPoints(29.7f);
                        gen = wordApp.CentimetersToPoints(21f);

                        if (Math.Round(wordApp.PointsToCentimeters(height), 2) != Math.Round(wordApp.PointsToCentimeters(yuk), 2))
                        {
                            cikti += i + ". SAYFA YÜKSEKLİK: " + Math.Round(wordApp.PointsToCentimeters(height), 3).ToString() + " ANCAK " + wordApp.PointsToCentimeters(yuk).ToString() + " olmalıdır!" + "\n";
                        }

                        if (Math.Round(wordApp.PointsToCentimeters(width), 2) != wordApp.PointsToCentimeters(gen))
                        {
                            cikti += i + ". SAYFA GENİŞLİK: " + Math.Round(wordApp.PointsToCentimeters(width), 2).ToString() + " ANCAK " + wordApp.PointsToCentimeters(gen).ToString() + " olmalıdır!" + "\n";
                        }


                        if (Math.Round(wordApp.PointsToCentimeters(topMargin), 2) != wordApp.PointsToCentimeters(ustKenarBoslugu))
                        {
                            cikti += i + ". SAYFA ÜST Kenar Boşluğunuz: " + Math.Round(wordApp.PointsToCentimeters(topMargin), 2).ToString() + " ANCAK " + wordApp.PointsToCentimeters(ustKenarBoslugu).ToString() + " olmalıdır!" + "\n";
                        }

                        if (Math.Round(wordApp.PointsToCentimeters(bottomMargin), 2) != wordApp.PointsToCentimeters(altKenarBoslugu))
                        {
                            cikti += i + ". SAYFA ALT Kenar Boşluğunuz: " + Math.Round(wordApp.PointsToCentimeters(bottomMargin), 2).ToString() + " ANCAK " + wordApp.PointsToCentimeters(altKenarBoslugu).ToString() + " olmalıdır!" + "\n";
                        }

                        if (Math.Round(wordApp.PointsToCentimeters(leftMargin), 2) != wordApp.PointsToCentimeters(solKenarBoslugu))
                        {
                            cikti += i + ". SAYFA SOL Kenar Boşluğunuz: " + Math.Round(wordApp.PointsToCentimeters(leftMargin), 2).ToString() + " ANCAK " + wordApp.PointsToCentimeters(solKenarBoslugu).ToString() + " olmalıdır!" + "\n";
                        }

                        if (Math.Round(wordApp.PointsToCentimeters(rightMargin), 2) != wordApp.PointsToCentimeters(sagKenarBoslugu))
                        {
                            cikti += i + ". SAYFA SAĞ Kenar Boşluğunuz: " + Math.Round(wordApp.PointsToCentimeters(rightMargin), 2).ToString() + " ANCAK " + wordApp.PointsToCentimeters(sagKenarBoslugu).ToString() + " olmalıdır!" + "\n";

                        }
                    }
                    else if (orient == "wdOrientLandscape")
                    {
                        ustKenarBoslugu = wordApp.CentimetersToPoints(3.25f);
                        altKenarBoslugu = wordApp.CentimetersToPoints(2.5f);
                        solKenarBoslugu = wordApp.CentimetersToPoints(2.5f);
                        sagKenarBoslugu = wordApp.CentimetersToPoints(3.0f);
                        yuk = wordApp.CentimetersToPoints(21f);
                        gen = wordApp.CentimetersToPoints(29.7f);

                        if (Math.Round(wordApp.PointsToCentimeters(height), 2) != Math.Round(wordApp.PointsToCentimeters(yuk), 2))
                        {
                            cikti += i + ". SAYFA YÜKSEKLİK: " + Math.Round(wordApp.PointsToCentimeters(height), 1).ToString() + " ANCAK " + wordApp.PointsToCentimeters(yuk).ToString() + " olmalıdır!" + "\n";
                        }

                        if (Math.Round(wordApp.PointsToCentimeters(width), 2) != Math.Round(wordApp.PointsToCentimeters(gen), 2))
                        {
                            cikti += i + ". SAYFA GENİŞLİK: " + Math.Round(wordApp.PointsToCentimeters(width), 2).ToString() + " ANCAK " + wordApp.PointsToCentimeters(gen).ToString() + " olmalıdır!" + "\n";
                        }

                        if (Math.Round(wordApp.PointsToCentimeters(topMargin), 2) != wordApp.PointsToCentimeters(ustKenarBoslugu))
                        {
                            cikti += i + ". SAYFA ÜST Kenar Boşluğunuz: " + Math.Round(wordApp.PointsToCentimeters(topMargin), 2).ToString() + " ANCAK " + wordApp.PointsToCentimeters(ustKenarBoslugu).ToString() + " olmalıdır!" + "\n";
                        }

                        if (Math.Round(wordApp.PointsToCentimeters(bottomMargin), 2) != wordApp.PointsToCentimeters(altKenarBoslugu))
                        {
                            cikti += i + ". SAYFA ALT Kenar Boşluğunuz: " + Math.Round(wordApp.PointsToCentimeters(bottomMargin), 2).ToString() + " ANCAK " + wordApp.PointsToCentimeters(altKenarBoslugu).ToString() + " olmalıdır!" + "\n";
                        }

                        if (Math.Round(wordApp.PointsToCentimeters(leftMargin), 2) != wordApp.PointsToCentimeters(solKenarBoslugu))
                        {
                            cikti += i + ". SAYFA SOL Kenar Boşluğunuz: " + Math.Round(wordApp.PointsToCentimeters(leftMargin), 2).ToString() + " ANCAK " + wordApp.PointsToCentimeters(solKenarBoslugu).ToString() + " olmalıdır!" + "\n";
                        }

                        if (Math.Round(wordApp.PointsToCentimeters(rightMargin), 2) != wordApp.PointsToCentimeters(sagKenarBoslugu))
                        {
                            cikti += i + ". SAYFA SAĞ Kenar Boşluğunuz: " + Math.Round(wordApp.PointsToCentimeters(rightMargin), 2).ToString() + " ANCAK " + wordApp.PointsToCentimeters(sagKenarBoslugu).ToString() + " olmalıdır!" + "\n";
                        }
                    }
                }
            }
            MessageBox.Show(cikti);
        }
    }
}
