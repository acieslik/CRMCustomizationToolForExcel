using DynamicsCRMCustomizationToolForExcel.Model;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DynamicsCRMCustomizationToolForExcel.Controller
{
    public class ComponentsTreeHandler
    {
        public delegate void CreateNewFormSheetDelegate(Guid formId);
        public delegate void CreateNewViewSheetDelegate(Guid viewId, Guid enityId);
        public delegate void CreateNewAttributesSheetDelegate(Guid enityId);
        private const string FORM_NODE_NAME = "Forms";
        private const string VIEW_NODE_NAME = "Views";
        private const string ATTRIBUTES_NODE_NAME = "Attributes";
        private TreeView componentsTree;
        private CreateNewFormSheetDelegate CreateNewFormSheetMethod;
        private CreateNewViewSheetDelegate CreateNewViewSheetMethod;
        private CreateNewAttributesSheetDelegate CreateNewAttributesSheetMethod;
        private ContextMenuStrip ctxMenuViews;
        private ContextMenuStrip ctxMenuViewItem;

        public ComponentsTreeHandler(TreeView componentsTree, EntityMetadata[] entMetadata, CreateNewFormSheetDelegate CreateNewFormSheetMethod, CreateNewViewSheetDelegate CreateNewViewSheetMethod, CreateNewAttributesSheetDelegate CreateNewAttributesSheetMethod)
        {
            this.componentsTree = componentsTree;
            this.CreateNewFormSheetMethod = CreateNewFormSheetMethod;
            this.CreateNewViewSheetMethod = CreateNewViewSheetMethod;
            this.CreateNewAttributesSheetMethod = CreateNewAttributesSheetMethod;
            componentsTree.NodeMouseClick += new TreeNodeMouseClickEventHandler(NodeMouseClick);
        }

        public void refreshTree()
        {
            componentsTree.Nodes.Clear();
            AddTreeViewData();
        }

        private void AddTreeViewData()
        {
            foreach (EntityMetadata item in GlobalApplicationData.Instance.currentEnitiesList)
            {
                string entityNode = string.Format("{0}  - ({1})", Utils.getLocalizedLabel(item.DisplayName.LocalizedLabels, GlobalApplicationData.Instance.currentLanguage), item.SchemaName);
                TreeNode node = new TreeNode(entityNode);
                node.Name = item.MetadataId.ToString();
                node.Nodes.Add(TreeViewEntityElements(item));
                node.Nodes.Add(TreeFormElement(item));
                node.Nodes.Add(TreeViewElement(item));
                componentsTree.Nodes.Add(node);
            }
        }

        private TreeNode TreeViewEntityElements(EntityMetadata entity)
        {
            TreeNode node = new TreeNode(ATTRIBUTES_NODE_NAME);
            node.Name = entity.MetadataId.ToString();
            return node;
        }

        private TreeNode TreeFormElement(EntityMetadata entity)
        {
            TreeNode node = new TreeNode(FORM_NODE_NAME);
            node.Name = entity.ObjectTypeCode.ToString();
            node.Nodes.Add("Temp","Temp");
            return node;
        }

        private TreeNode TreeViewElement(EntityMetadata entity)
        {
            TreeNode node = new TreeNode(VIEW_NODE_NAME);
            node.Name = entity.ObjectTypeCode.ToString();
            node.Nodes.Add("Temp", "Temp");
            return node;
        }

        #region Event Handler

        private void ExpandFormNode(TreeNodeMouseClickEventArgs e)
        {
            if (e.Node.Nodes.ContainsKey("Temp"))
                e.Node.Nodes.RemoveByKey("Temp");
            if (e.Node.Nodes.Count == 0)
            {
                int typecode;
                if (int.TryParse(e.Node.Name, out typecode))
                {
                    EntityCollection forms = GlobalOperations.Instance.CRMOpHelper.GetForms(typecode);
                    foreach (Entity item in forms.Entities)
                    {
                        if (item.Contains("name"))
                        {
                            TreeNode node = new TreeNode(item.Attributes["name"].ToString());
                            node.Name = item.Id.ToString();
                            e.Node.Nodes.Add(node);
                        }
                    }
                    e.Node.Expand();
                }
            }
        }

        private void ExpandViewNode(TreeNodeMouseClickEventArgs e)
        {
            if (e.Node.Nodes.ContainsKey("Temp"))
                e.Node.Nodes.RemoveByKey("Temp");
            if (e.Node.Nodes.Count == 0)
            {
                int typecode;
                if (int.TryParse(e.Node.Name, out typecode))
                {
                    EntityCollection views = GlobalOperations.Instance.CRMOpHelper.GetViews(typecode);
                    foreach (Entity item in views.Entities)
                    {
                        if (item.Contains("name"))
                        {
                            TreeNode node = new TreeNode(item.Attributes["name"].ToString());
                            node.Name = item.Id.ToString();
                            e.Node.Nodes.Add(node);
                        }
                    }
                    e.Node.Expand();
                }
            }
        }

        private void CreateNewFormSheet(TreeNodeMouseClickEventArgs e)
        {
            Guid formGuid;
            if (Guid.TryParse(e.Node.Name, out formGuid))
            {
                CreateNewFormSheetMethod(formGuid);
            }
        }

        private void CreateNewViewSheet(TreeNodeMouseClickEventArgs e)
        {
            Guid viewGuid, enityGuid;
            if (Guid.TryParse(e.Node.Name, out viewGuid) && Guid.TryParse(e.Node.Parent.Parent.Name, out enityGuid))
            {
                CreateNewViewSheetMethod(viewGuid, enityGuid);
            }
        }

        private void CreateNewAttributesSheet(TreeNodeMouseClickEventArgs e)
        {
            Guid entityGuid;
            if (Guid.TryParse(e.Node.Name, out entityGuid))
            {
                CreateNewAttributesSheetMethod(entityGuid);
            }
        }

        private void NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                if (e.Node.Text == FORM_NODE_NAME)
                {
                    ExpandFormNode(e);
                }
                if (e.Node.Text == VIEW_NODE_NAME)
                {
                    ExpandViewNode(e);
                }
                else if (e.Node.Parent != null && e.Node.Parent.Text == FORM_NODE_NAME)
                {
                    CreateNewFormSheet(e);
                }
                else if (e.Node.Parent != null && e.Node.Parent.Text == VIEW_NODE_NAME)
                {
                    CreateNewViewSheet(e);
                }
                else if (e.Node.Text == ATTRIBUTES_NODE_NAME)
                {
                    CreateNewAttributesSheet(e);
                }
            }
        }


        public void addTreeViewContextMenuForViews(ContextMenuStrip ctxMenuViews, ContextMenuStrip ctxMenuViewItem)
        {
            componentsTree.MouseUp += new MouseEventHandler(componentsTree_MouseUp);
            this.ctxMenuViews = ctxMenuViews;
            this.ctxMenuViewItem = ctxMenuViewItem;
        }

        private void componentsTree_MouseUp(object sender, MouseEventArgs e)
        {
            //if (e.Button == MouseButtons.Right)
            //{
            //    // Select the clicked node
            //    componentsTree.SelectedNode = componentsTree.GetNodeAt(e.X, e.Y);

            //    if (componentsTree.SelectedNode != null)
            //    {
            //        if (componentsTree.SelectedNode.Text == VIEW_NODE_NAME)
            //        {
            //            Guid resultguid;
            //            string name = componentsTree.SelectedNode.Parent.Name;
            //            if (Guid.TryParse(name, out resultguid))
            //            {
            //                GlobalApplicationData.Instance.selectedEntityTree = resultguid;
            //                ctxMenuViews.Show(componentsTree, e.Location);
            //            }
            //        }
            //        else if (componentsTree.SelectedNode.Parent != null && componentsTree.SelectedNode.Parent.Text == VIEW_NODE_NAME)
            //        {
            //            Guid resultguid;
            //            string name = componentsTree.SelectedNode.Name;
            //            if (Guid.TryParse(name, out resultguid))
            //            {
            //                GlobalApplicationData.Instance.selectedEntityViewTree = resultguid;
            //                ctxMenuViewItem.Show(componentsTree, e.Location);
            //            }
            //        }
            //    }
            //}
        }
        #endregion
    }
}