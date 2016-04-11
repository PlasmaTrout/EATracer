using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Tools;

namespace EAWordTracer
{
    public partial class ThisAddIn
    {

        EASelectionControl control;
        Microsoft.Office.Tools.CustomTaskPane pane;

        public CustomTaskPane TaskPane
        {
            get
            {
                return pane;
            }
        }
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            control = new EASelectionControl();
            pane = this.CustomTaskPanes.Add(control, "EA Tracability Report");
            pane.Width = 450;

            pane.Visible = true;
            pane.VisibleChanged += Pane_VisibleChanged;
        }

        private void Pane_VisibleChanged(object sender, EventArgs e)
        {
            Globals.Ribbons.EARibbon.tracerToggleButton.Checked = control.Visible;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
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
