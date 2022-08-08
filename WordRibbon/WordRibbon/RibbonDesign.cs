using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WordRibbon
{
    public partial class RibbonDesign
    {
        private void RibbonDesign_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btn_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("Merhaba!");
        }
    }
}
