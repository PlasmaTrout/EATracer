using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace EAAffectedElementsReport
{
    public partial class EAWordRibbon
    {
        private void EAWordRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void launchTReportButton_Click(object sender, RibbonControlEventArgs e)
        {
            EASelectionControl control = new EASelectionControl();
            control.Show();
        }
    }
}
