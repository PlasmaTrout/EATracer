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
                    repos.Exit();
                }

                statusLabel.Text = "Opening Repository...";
                statusStrip1.Refresh();
                var opened = repos.OpenFile(path);
                if (opened)
                {
                    EA.Collection coll = repos.Models;
                    statusLabel.Text = "Loading Models...";
                    statusStrip1.Refresh();
                    LoadModelCombo(coll);
                }
                else
                {

                }
            }
            catch (Exception e)
            {
                repos.CloseFile();
            }
        }

        private void LoadModelCombo(EA.Collection coll)
        {
            IEnumerator e = coll.GetEnumerator();
            while (e.MoveNext())
            {
                EA.Package element = (EA.Package)e.Current;
                modelCombo.Items.Add(element.Name);
            }
            statusLabel.Text = "Select Model";

        }

        private void RecursePackages(EA.Collection packages)
        {
            IEnumerator e = packages.GetEnumerator();

            while (e.MoveNext())
            {
                EA.Package package = (EA.Package)e.Current;
                comboPackage.Items.Add($"{package.Name} | {package.PackageID}");
                if (package.Packages.Count > 0)
                {
                    RecursePackages(package.Packages);
                }
            }

            
        }

        private void LoadElements(EA.Collection elements)
        {
            IEnumerator e = elements.GetEnumerator();

            while (e.MoveNext())
            {
                EA.Element element = (EA.Element)e.Current;
                comboElements.Items.Add($"{element.Name} | {element.ElementID}");
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

        private void comboElements_SelectedIndexChanged(object sender, EventArgs e)
        {
            string element = comboElements.SelectedItem.ToString();
            string id = element.Split('|')[1].Trim();

            EA.Element root = this.Repository.GetElementByID(int.Parse(id));

            GetConnections(root, "forward");
            AddHeaders();
        }

        private void modelCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            String value = modelCombo.SelectedItem.ToString();
            statusLabel.Text = "Getting Packages ...";
            EA.Package package = (EA.Package)repos.Models.GetByName(value);
            RecursePackages(package.Packages);
            statusLabel.Text = "Select Package!";
        }

        #endregion

        private void comboPackage_SelectedIndexChanged(object sender, EventArgs e)
        {
            String value = comboPackage.SelectedItem.ToString();
            String pkid = value.Split('|')[1].Trim();
            statusLabel.Text = "Getting Elements...";
            EA.Package package = (EA.Package)repos.GetPackageByID(int.Parse(pkid));
            LoadElements(package.Elements);

            statusLabel.Text = "Select Element!";
        }
    }
}
