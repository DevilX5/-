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
                toolStripLabel1.Text = $"文件夹路径：{CurrentPath}";
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
                tn.Nodes.Add("正在读取...");
                nodelst.Add(tn);
            }
            treeView1.Nodes.AddRange(nodelst.ToArray());
        }
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            var dt = new DataTable();
            toolStripLabel1.Text = "正在合并数据...";
            var s = treeView1.AllTreeNodes().Where(n => !n.Text.Equals("正在读取...") &&n.Checked).Select(n => n.FullPath);
            var c =s.Select(n => n.Split('\\')[0]).Distinct().Select(n=>new FileAndSheets() {FileName=n}).ToList();
            var tklst = new List<Task>();
            var ds = new DataSet("ds");
            c.ForEach(f =>
            {
                f.SheetNames = s.Where(n => n.Contains(f.FileName) && n.Contains("\\")).Select(n => n.Split('\\')[1]).ToList();
                var t = Task.Run(() =>
                {
                    f.FileName = $"{CurrentPath}\\{f.FileName}";
                    if (f.SheetNames.Count == 0)
                        f.SheetNames = ExcelHelper.GetSheetList(f.FileName);
                    ExcelHelper.SetDataSet(ds, f);
                });
                var waittk = Task.Run(() =>
                {
                    Task.WaitAll(t);
                });
                tklst.Add(waittk);
            });
            Task.Run(() =>
            {
                Task.WaitAll(tklst.ToArray());
                foreach (DataTable sdt in ds.Tables)
                {
                    dt.Merge(sdt);
                }
                this.Invoke((MethodInvoker)delegate
                {
                    dataGridView1.DataSource = dt;
                    toolStripLabel1.Text = "数据合并完毕";
                });
            });
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

        private void 合并文件夹_Load(object sender, EventArgs e)
        {

        }

        private void treeView1_AfterExpand(object sender, TreeViewEventArgs e)
        {
            var currentNode = e.Node;
            var filepath = $"{CurrentPath}\\{currentNode.FullPath}";
            currentNode.Nodes.Clear();
            currentNode.Nodes.Add("正在读取...");
            var t = Task.Run(() =>
            {
                if (Path.GetExtension(filepath).Equals(".xlsx"))
                {
                    return EppHelper.GetAllSheet(filepath);
                }
                else
                {
                    return NpoiHelper.GetAllSheet(filepath);
                }
            });
            Task.Run(() =>
            {
                Task.WaitAll(t);
                this.Invoke((MethodInvoker)delegate
                {
                    currentNode.Nodes.Clear();
                    t.Result.ForEach(f => 
                    {
                        currentNode.Nodes.Add(f);
                    });
                });
            });
        }

        private void dataGridView1_RowStateChanged(object sender, DataGridViewRowStateChangedEventArgs e)
        {
            e.Row.HeaderCell.Value = string.Format("{0}", e.Row.Index + 1);
        }
    }
}