using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 合并表
{
    public partial class 合并文件夹 : Form
    {
        public 合并文件夹()
        {
            InitializeComponent();
        }
        public string CurrentPath { get; set; }
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择文件路径";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                CurrentPath = dialog.SelectedPath;
                GetFolderFiles();
            }
        }
        void GetFolderFiles()
        {
            DirectoryInfo folder = new DirectoryInfo(CurrentPath);
            var nodelst = new List<TreeNode>();
            foreach (FileInfo file in folder.GetFiles("*.xls*"))
            {
                var tn = new TreeNode();
                tn.Text = file.Name;
                tn.Nodes.Add("");
                nodelst.Add(tn);
            }
            treeView1.Nodes.AddRange(nodelst.ToArray());
        }
        private void toolStripButton2_Click(object sender, EventArgs e)
        {

        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            SaveFileDialog kk = new SaveFileDialog();
            kk.Title = "保存EXCEL文件";
            kk.Filter = "excel文件(*.xlsx) |*.xlsx ";
            kk.FilterIndex = 1;
            if (kk.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    toolStripLabel1.Text = "正在导出数据";
                    var dt = dataGridView1.DataSource as DataTable;
                    var t = Task.Run(() =>
                    {
                        return EppHelper.ExportByDt(kk.FileName, dt);
                    });
                    Task.Run(() =>
                    {
                        Task.WaitAll(t);
                        if (t.Result)
                        {
                            this.Invoke((MethodInvoker)delegate
                            {
                                toolStripLabel1.Text = "导出成功";
                            });
                        }
                    });

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            SetCheckState(treeView1.Nodes, true);
        }
        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            SetCheckState(treeView1.Nodes, false);
        }
        void SetCheckState(TreeNodeCollection tc,bool selectall)
        {
            foreach (TreeNode item in tc)
            {
                item.Checked = selectall ? true : !item.Checked;
                if (item.Nodes.Count > 0)
                {
                    SetCheckState(item.Nodes, selectall);
                }
            }
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {

        }

        private void 合并文件夹_Load(object sender, EventArgs e)
        {
            CurrentPath = "C:\\Users\\DX\\Desktop\\新建文件夹";
            GetFolderFiles();
        }

        private void treeView1_BeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            MessageBox.Show(treeView1.SelectedNode.FullPath);
            //if(treeView1.SelectedNode.FullPath)
            //var file = $"{CurrentPath}\\{treeView1.SelectedNode.Text}";
            //if (Path.GetExtension(CurrentPath) == "xlsx")
            //{
            //     EppHelper.GetAllSheet(file);
            //}
            //else
            //{
            //    NpoiHelper.GetAllSheet(file);
            //}

        }

        private void treeView1_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            treeView1.SelectedNode.Expand();
        }
    }
}