using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections;
using Microsoft.Office.Interop.Word;

namespace EAWordTracer
{
    public partial class EASelectionControl : UserControl
    {
        EA.Repository repos;
        
        public Table CurrentTable { get; set; }

        public EA.Repository Repository
        {
            get
            {
                return repos;
            }
        }
        public EASelectionControl()
        {
            InitializeComponent();
        }

        private void OpenRepo(string path)
        {
            try
            {
                if (repos == null)
                {
                    repos = new EA.Repository();
                }else
                {
                    repos.CloseFile();
                    projectTree.Nodes.Clear();
                }

                statusLabel.Text = "Opening Repository...";
                statusStrip1.Refresh();
                var opened = repos.OpenFile(path);
                if (opened)
                {
                    statusLabel.Text = "Loading Tree...";
                    statusStrip1.Refresh();
                    var loader = new TreeLoader(this.Repository);
                    loader.NodeAdded += Loader_NodeAdded;
                    TreeNode root = loader.BuildTree();
                    projectTree.Nodes.Add(root);
                    statusLabel.Text = "Select Node!";         
                }
            }
            catch (Exception)
            {
                repos.Exit();
            }
        }

        private void Loader_NodeAdded(TreeNode label)
        {
            statusLabel.Text = "Loading Tree..." + label.Text;
            statusStrip1.Refresh();
        }


        #region Events
        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog();

            if (result == DialogResult.OK)
            {
                fileTextBox.Text = openFileDialog1.FileName;
                OpenRepo(fileTextBox.Text);

            }
        }

        #endregion

        private void generateButton_Click(object sender, EventArgs e)
        {
            var selected = projectTree.SelectedNode;

            if(selected.ImageIndex > 1)
            {
                statusLabel.Text = "Getting Selected Element....";
                EA.Element element = this.Repository.GetElementByID(int.Parse(selected.Tag.ToString()));
                statusLabel.Text = "Running Report...";

                Range range = Globals.ThisAddIn.Application.Selection.Range;
                this.CurrentTable = Globals.ThisAddIn.Application.ActiveDocument.Tables.Add(range, 2, 2);
                this.CurrentTable.ApplyStyleHeadingRows = true;
                this.CurrentTable.Rows[1].Cells[1].Range.Text = "Directly Impacted (Dependent)";
                this.CurrentTable.Rows[1].Cells[2].Range.Text = "Possibly Impacted (Downstream)";
                //this.CurrentTable.set_Style("Table Grid 8");
                
                TableLoader loader = new TableLoader(this.Repository, this.CurrentTable);
                loader.ReportElementTraversed += Loader_ReportElementTraversed1; ;
                loader.RenderTable(element);

                statusLabel.Text = "Done!";
            }
        }

        private void Loader_ReportElementTraversed1(string element, string path, bool direct)
        {
            int cell = 1;

            if (!direct)
            {
                cell = 2;
            }

            statusLabel.Text = $"Running Report...{path}";

            this.CurrentTable.Rows[2].Cells[cell].Range.Text = this.CurrentTable.Rows[2].Cells[cell].Range.Text + $" {element}";
        }

    }
}
