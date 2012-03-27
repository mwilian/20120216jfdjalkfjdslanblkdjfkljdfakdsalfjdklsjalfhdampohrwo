using System;
using System.Data;
using System.Configuration;

using System.Web;




using System.ComponentModel;
using QueryBuilder;

namespace QueryDesigner
{
    public class RadTreeViewLoader
    {
        private static string SubTable = "S";
        public static RadTreeView LoadTree(ref RadTreeView result, BindingList<Node> list, string rootCode, string fileLanguage)
        {
            RadTreeNode root = new RadTreeNode();
            root.Name = rootCode;
            root.Text = rootCode;

            //    Configuration.clsConfigurarion config = new QueryDesigner.Configuration.clsConfigurarion();
            //    config.GetDictionary(GetServerDirPath() + "/Configuration/Languages/" + fileLanguage);
            for (int i = 0; i < list.Count; i++)
            {
                Node node = list[i];
                //   node.Description = config.GetValueDictionary(node.Description);
                // RadTreeNode radNode = GetRadTreeNode(node);
                InsertRadTreeNode(ref root, node);
            }
            result.Nodes.Clear();
            while (root.Nodes.Count > 0)
            {
                result.Nodes.Add(root.Nodes[0]);
            }
            return result;
        }
        //private static string GetServerDirPath()
        //{
        //    return "http://" + HttpContext.Current.Request.Url.Authority;
        //}
        private static void InsertRadTreeNode(ref RadTreeNode desRadTreeNode, Node srcNode)
        {
            if (desRadTreeNode.Name == Node.GetFamily(srcNode.Code))
            {
                desRadTreeNode.Nodes.Add(GetRadTreeNode(srcNode));
                return;
            }
            else
            {
                if (desRadTreeNode.Nodes.Count > 0)
                {
                    desRadTreeNode.Image = Properties.Resources.Folder;
                    RadTreeNode radTreeNodeParent;
                    RadTreeNode[] arr_node = desRadTreeNode.Nodes.Find(Node.GetFamily(srcNode.Code), false);
                    if (arr_node.Length > 0)
                    {
                        radTreeNodeParent = arr_node[0];
                    }
                    else
                        radTreeNodeParent = null;
                    if (radTreeNodeParent != null)
                    {
                        radTreeNodeParent.Image = Properties.Resources.Folder;
                        radTreeNodeParent.Nodes.Add(GetRadTreeNode(srcNode));
                        return;
                    }
                    else
                    {
                        for (int i = desRadTreeNode.Nodes.Count - 1; i >= 0; i--)
                        {
                            RadTreeNode tmp = desRadTreeNode.Nodes[i];
                            InsertRadTreeNode(ref tmp, srcNode);
                        }
                    }
                }
            }
        }

        private static RadTreeNode GetRadTreeNode(Node node)
        {
            RadTreeNode result = new RadTreeNode();
            result.AllowDrop = false;
            result.Name = node.Code;


            result.Text = node.Description;
            result.Tag = node.Agregate + ";" +
                                node.FType + ";" + node.NodeDesc;
            result.Image = Properties.Resources.Field;
            return result;
        }
    }
}
