using System;
using System.Data;
using System.Configuration;

using System.Web;




using System.ComponentModel;
using QueryBuilder;
using System.Windows.Forms;

namespace dCube
{
    public class TreeViewLoader
    {
        private static string SubTable = "S";
        public static TreeView LoadTree(ref TreeView result, BindingList<Node> list, string rootCode, string fileLanguage)
        {
            TreeNode root = new TreeNode();
            root.Name = rootCode;
            root.Text = rootCode;

            //    Configuration.clsConfigurarion config = new dCube.Configuration.clsConfigurarion();
            //    config.GetDictionary(GetServerDirPath() + "/Configuration/Languages/" + fileLanguage);
            for (int i = 0; i < list.Count; i++)
            {
                Node node = list[i];
                //   node.Description = config.GetValueDictionary(node.Description);
                // TreeNode radNode = GetTreeNode(node);
                InsertTreeNode(ref root, node);
            }
            result.Nodes.Clear();
            for (int i = 0; i < root.Nodes.Count; i++)
            {
                result.Nodes.Add(root.Nodes[i]);
            }
            return result;
        }
        //private static string GetServerDirPath()
        //{
        //    return "http://" + HttpContext.Current.Request.Url.Authority;
        //}
        private static void InsertTreeNode(ref TreeNode desTreeNode, Node srcNode)
        {
            if (desTreeNode.Name == Node.GetFamily(srcNode.Code))
            {
                desTreeNode.Nodes.Add(GetTreeNode(srcNode));
                return;
            }
            else
            {
                if (desTreeNode.Nodes.Count > 0)
                {
                    desTreeNode.SelectedImageKey = desTreeNode.ImageKey = "Folder";

                    TreeNode TreeNodeParent;
                    TreeNode[] arr_node = desTreeNode.Nodes.Find(Node.GetFamily(srcNode.Code), false);
                    if (arr_node.Length > 0)
                    {
                        TreeNodeParent = arr_node[0];
                    }
                    else
                        TreeNodeParent = null;
                    if (TreeNodeParent != null)
                    {
                        //TreeNodeParent.SelectedImageKey = TreeNodeParent.ImageKey = "Folder";
                        TreeNodeParent.Nodes.Add(GetTreeNode(srcNode));
                        return;
                    }
                    else
                    {
                        for (int i = desTreeNode.Nodes.Count - 1; i >= 0; i--)
                        {
                            TreeNode tmp = desTreeNode.Nodes[i];
                            InsertTreeNode(ref tmp, srcNode);
                        }
                    }
                }
            }
        }

        private static TreeNode GetTreeNode(Node node)
        {
            TreeNode result = new TreeNode();
            //result.A = false;
            result.Name = node.Code;


            result.Text = node.Description;
            result.Tag = node;
            //result.SelectedImageKey = result.ImageKey = "Field";
            return result;
        }
    }
}
