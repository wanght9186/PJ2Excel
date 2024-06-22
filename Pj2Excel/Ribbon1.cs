using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Pj2Excel
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void ToExcel_Click(object sender, RibbonControlEventArgs e)
        {
            TransToExcelFuc.P2J_TransSelectionToExcel();
        }
    }
}
