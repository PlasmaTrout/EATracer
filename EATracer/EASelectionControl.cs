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
using Microsoft.Office.Interop.Excel;

namespace EATracer
{
    public partial class EASelectionControl : UserControl
    {
        EA.Repository repos;
        Dictionary<string, int> completionDictionary = new Dictionary<string, int>();
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


        private void GetConnections(EA.Element element, string traversal)
        {
            EA.Collection connectors = element.Connectors;

            for (short i = 0; i < connectors.Count; i++)
            {
                EA.Connector connector = connectors.GetAt(i);
                EA.Element target = this.Repository.GetElementByID(connector.SupplierID);
                EA.Element source = this.Repository.GetElementByID(connector.ClientID);
                String key = target.Name + ":" + source.Name;

                if (completionDictionary.ContainsKey(key))
                {
                    completionDictionary[key] = completionDictionary[key] + 1;
                }
                else
                {
                    completionDictionary.Add(key, 1);
                    AddRow(source, target, traversal, connector);
                    GetConnections(target, "forwards");
                    GetConnections(source, "backwards");
                }
            }
        }

        private void AddRow(EA.Element source, EA.Element destination, string traversal, EA.Connector connector)
        {
            var sheet = Globals.ThisAddIn.Application.ActiveSheet;
            sheet.Range["A1"].EntireRow.Insert(XlInsertShiftDirection.xlShiftDown);
            sheet.Range["A1"].Value = source.Name;
            sheet.Range["B1"].Value = source.Type;
            sheet.Range["C1"].Value = destination.Name;
            sheet.Range["D1"].Value = destination.Type;
            sheet.Range["E1"].Value = traversal;
            sheet.Range["F1"].Value = connector.Type;
            sheet.Range["G1"].Value = connector.Stereotype;
            sheet.Range["H1"].Value = connector.Direction;
           
        }

        private void AddHeaders()
        {
            var sheet = Globals.ThisAddIn.Application.ActiveSheet;
            sheet.Range["A1"].EntireRow.Insert(XlInsertShiftDirection.xlShiftDown);
            sheet.Range["A1"].Value = "Source";
            sheet.Range["B1"].Value = "Source Type";
            sheet.Range["C1"].Value = "Destination";
            sheet.Range["D1"].Value = "Dest Type";
            sheet.Range["E1"].Value = "Traversal";
            sheet.Range["F1"].Value = "Connector";
            sheet.Range["G1"].Value = "StereoType";
            sheet.Range["H1"].Value = "Direction";

            
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
                GetConnections(element, "forward");
                AddHeaders();
                statusLabel.Text = "Done!";
            }
        }
    }
}
