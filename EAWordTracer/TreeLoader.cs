using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EAWordTracer
{
    public delegate void NodeAddedHandler(TreeNode label);

    public class TreeLoader
    {

        public event NodeAddedHandler NodeAdded;

        public EA.Repository Repository { get; set; }

        public TreeLoader(EA.Repository repository)
        {
            this.Repository = repository;
        }

        public TreeNode BuildTree()
        {
            TreeNode root = new TreeNode("Project");
            LoadModels(root);
            return root;
        }

        private void LoadModels(TreeNode current)
        {
            var e = this.Repository.Models.GetEnumerator();
            while (e.MoveNext())
            {
                TreeNode model = new TreeNode();
                EA.Package pkg = (EA.Package)e.Current;
                model.Text = pkg.Name;
                model.Tag = pkg.PackageID;
                model.ImageIndex = 0;
                model.SelectedImageIndex = 1;
                current.Nodes.Add(model);

                NodeAdded?.Invoke(model);

                LoadPackages(pkg, model);

               
                
            }
        }

        private void LoadPackages(EA.Package package, TreeNode current)
        {
            var e = package.Packages.GetEnumerator();

            while (e.MoveNext())
            {
                EA.Package pkg = (EA.Package)e.Current;
                TreeNode newNode = new TreeNode();
                newNode.Text = pkg.Name;
                newNode.Tag = pkg.PackageID;
                newNode.ImageIndex = 0;
                newNode.SelectedImageIndex = 1;

                current.Nodes.Add(newNode);
                NodeAdded?.Invoke(newNode);

                if(pkg.Packages.Count > 0)
                {
                    LoadPackages(pkg, newNode);
                }

                if(pkg.Elements.Count > 0)
                {
                    LoadElements(pkg.Elements, newNode);
                }
            }
        }

        private void LoadElements(EA.Collection elements, TreeNode current)
        {
            var e = elements.GetEnumerator();

            while (e.MoveNext())
            {
                EA.Element element = (EA.Element)e.Current;
                TreeNode eNode = new TreeNode();
                eNode.Text = element.Name;
                eNode.Tag = element.ElementID;
                eNode.ImageIndex = 2;
                eNode.SelectedImageIndex = 3;

                current.Nodes.Add(eNode);
                NodeAdded?.Invoke(eNode);

                if (element.Elements.Count > 0)
                {
                    LoadElements(element.Elements, eNode);
                }
            }
        }
    }
}
