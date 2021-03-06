﻿using System;
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

       
        private TableLoadDirection GetDirection()
        {
            if (radioBackwards.Checked)
            {
                return TableLoadDirection.Backwards;
            }
            if (radioForwards.Checked)
            {
                return TableLoadDirection.Forwards;
            }
            return TableLoadDirection.Both;
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

                TableLoader loader = new TableLoader(this.Repository,this.GetDirection());
                loader.ReportElementTraversed += Loader_ReportElementTraversed;
                loader.RenderTable(element);

                statusLabel.Text = "Done!";
            }
        }

        private void Loader_ReportElementTraversed(string element)
        {
            statusLabel.Text = "Running Report..." + element;
            statusStrip1.Refresh();
        }
    }
}
