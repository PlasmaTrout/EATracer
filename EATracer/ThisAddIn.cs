using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace EATracer
{
    public partial class ThisAddIn
    {
        EASelectionControl control;
        Microsoft.Office.Tools.CustomTaskPane pane;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            control = new EASelectionControl();
            pane = this.CustomTaskPanes.Add(control, "EA Tracer");
            pane.Visible = true;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            pane.Visible = false;
            control.Repository.CloseFile();
            control.Repository.Exit();
            this.CustomTaskPanes.Remove(pane);
            control = null;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
