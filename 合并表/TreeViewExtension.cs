using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 合并表
{
    public static class TreeViewExtension
    {
        public static IEnumerable<TreeNode> AllTreeNodes(this TreeView treeView)
        {
            Queue<TreeNode> nodes = new Queue<TreeNode>();
            foreach (TreeNode item in treeView.Nodes)
                nodes.Enqueue(item);

            while (nodes.Count > 0)
            {
                TreeNode node = nodes.Dequeue();
                yield return node;
                foreach (TreeNode item in node.Nodes)
                    nodes.Enqueue(item);
            }
        }
    }
}
