using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;



namespace Ribbon
{
    public partial class myRibbon
    {
        private void Form_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Kontroller kontroller = new Kontroller();
            //kontroller.BaslikVarMi("ÖNSÖZ", "Times New Roman", 16, -1, -1, 24, WdLineSpacing.wdLineSpaceSingle, WdParagraphAlignment.wdAlignParagraphCenter);
            //kontroller.BaslikVarMi("İçindekiler", "Times New Roman", 16, -1, -1, 24, WdLineSpacing.wdLineSpaceSingle, WdParagraphAlignment.wdAlignParagraphCenter);
            //kontroller.BaslikVarMi("Özet", "Times New Roman", 16, -1, -1, 24, WdLineSpacing.wdLineSpaceSingle, WdParagraphAlignment.wdAlignParagraphCenter);
            //kontroller.BaslikVarMi("Abstract", "Times New Roman", 16, -1, -1, 24, WdLineSpacing.wdLineSpaceSingle, WdParagraphAlignment.wdAlignParagraphCenter);
            //kontroller.BaslikVarMi("Şekiller Listesi", "Times New Roman", 16, -1, -1, 24, WdLineSpacing.wdLineSpaceSingle, WdParagraphAlignment.wdAlignParagraphCenter);
            //kontroller.BaslikVarMi("Tablolar Listesi", "Times New Roman", 16, -1, -1, 24, WdLineSpacing.wdLineSpaceSingle, WdParagraphAlignment.wdAlignParagraphCenter);
            //kontroller.BaslikVarMi("Simgeler ve Kısaltmalar", "Times New Roman", 16, -1, -1, 24, WdLineSpacing.wdLineSpaceSingle, WdParagraphAlignment.wdAlignParagraphCenter);
             //kontroller.sayfaKenarBoslukveA4Kontrol();
             kontroller.onBolumBaslikkKontrol("Times New Roman", 16, -1, -1, 24, WdLineSpacing.wdLineSpaceSingle, WdParagraphAlignment.wdAlignParagraphCenter);
        }
    }
}
//  string textFromDoc = Globals.ThisAddIn.Application.ActiveDocument.Range(0, 20).Text;
// MessageBox.Show(textFromDoc);