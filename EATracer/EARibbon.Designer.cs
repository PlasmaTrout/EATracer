namespace EATracer
{
    partial class EARibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public EARibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.traceabilityGroup = this.Factory.CreateRibbonGroup();
            this.tracerToggleButton = this.Factory.CreateRibbonToggleButton();
            this.tab1.SuspendLayout();
            this.traceabilityGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.traceabilityGroup);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // traceabilityGroup
            // 
            this.traceabilityGroup.Items.Add(this.tracerToggleButton);
            this.traceabilityGroup.Label = "Traceability";
            this.traceabilityGroup.Name = "traceabilityGroup";
            // 
            // tracerToggleButton
            // 
            this.tracerToggleButton.Label = "EA Tracer Pane";
            this.tracerToggleButton.Name = "tracerToggleButton";
            this.tracerToggleButton.OfficeImageId = "BusinessFormWizard";
            this.tracerToggleButton.ShowImage = true;
            this.tracerToggleButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.tracerToggleButton_Click);
            // 
            // EARibbon
            // 
            this.Name = "EARibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.EARibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.traceabilityGroup.ResumeLayout(false);
            this.traceabilityGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup traceabilityGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton tracerToggleButton;
    }

    partial class ThisRibbonCollection
    {
        internal EARibbon EARibbon
        {
            get { return this.GetRibbon<EARibbon>(); }
        }
    }
}
