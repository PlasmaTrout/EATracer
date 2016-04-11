using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace EAWordTracer
{
    public partial class EARibbon
    {
        private void EARibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }


        private void tracerToggleButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.TaskPane.Visible = ((RibbonToggleButton)sender).Checked;
        }
    }
}
