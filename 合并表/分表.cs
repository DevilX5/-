using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 合并表
{
    public partial class 分表 : Form
    {
        string file;
        List<string> excelColumns;
        DataTable dtresult;
        List<string> headNameList = new List<string>();
        List<string> columnConfig = new List<string>();
        public string CurrentPath { get; set; }
        public 分表()
        {
            InitializeComponent();
        }

        private void 分表_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            file = files[0];
            CurrentPath = file;
            Text = CurrentPath;
            var t = Task.Run(() =>
            {
                return ExcelHelper.GetSheetList(CurrentPath);
            });
            Task.Run(() =>
            {
                Task.WaitAll(t);
                this.Invoke((MethodInvoker)delegate
                {
                    comboBox1.Items.Clear();
                    t.Result.ForEach(n => comboBox1.Items.Add(n));
                    comboBox1.SelectedIndex = 0;
                });
            });
            GC.Collect();
        }

        private void 分表_DragEnter(object sender, DragEventArgs e)
        {
            GC.Collect();
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void btnFind_Click(object sender, EventArgs e)
        {
            var fs = new FileAndSheets();
            fs.FileName = CurrentPath;
            fs.CurrentSheetName = comboBox1.Text;
            var t = Task.Run(() =>
            {
                if (Path.GetExtension(CurrentPath).Equals(".xlsx"))
                {
                    dtresult = EppHelper.GetSingleSheetData(fs);
                }
                else
                {
                    dtresult = NpoiHelper.GetSingleSheetData(fs);
                }
                return dtresult;
            });
            Task.Run(() => {
                Task.WaitAll(t);
                this.Invoke((MethodInvoker)delegate
                {
                    dataGridView1.DataSource = dtresult;
                    foreach (DataColumn dc in dtresult.Columns)
                    {
                        headNameList.Add(dc.ColumnName);
                    }
                    comboBox2.Items.Clear();
                    headNameList.ForEach(n => comboBox2.Items.Add(n));
                    comboBox2.SelectedIndex = 0;
                });
            });
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择文件路径";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string foldPath = dialog.SelectedPath;
                textBox2.Text = foldPath;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "选择模版";
            ofd.Filter = "所有文件(*.*)|*.*";
            ofd.FilterIndex = 1;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBox3.Text = ofd.FileName;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            columnConfig.Clear();
            string fileName = textBox1.Text.Trim().ToString();
            string columnName = comboBox2.Text;
            string filePath = textBox2.Text;
            string muban = textBox3.Text;
            if (string.IsNullOrEmpty(filePath) || string.IsNullOrEmpty(columnName) || string.IsNullOrEmpty(muban))
            {
                MessageBox.Show("请提供完整的保存信息");
            }
            else if (Path.GetExtension(muban).Equals(".xls")) MessageBox.Show("模板仅支持07以上版本Excel文件,文件后缀为.xlsx");
            else
            {
                columnConfig = dtresult.AsEnumerable().GroupBy(n =>new { name = n[columnName].ToString() }).Select(n => n.Key.name).ToList();
                var t = Task.Run(() =>
                {
                    columnConfig.ForEach(f =>
                    {
                        var sourceFile = muban;
                        var destinationFile = $"{filePath}/{f}{fileName}.xlsx";
                        File.Copy(sourceFile, destinationFile, true);
                    });
                });
                Task.Run(() => {
                    Task.WaitAll(t);
                    MessageBox.Show("模板创建完毕");
                });
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            columnConfig.Clear();
            string sheetName = textBox4.Text.Trim().ToString();
            string fileName = textBox1.Text.Trim().ToString();
            string columnName = comboBox2.Text;
            string filePath = textBox2.Text;
            string startCell = textBox5.Text.Trim().ToString();
            bool flag = checkBox1.Checked;
            if (string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(columnName) || string.IsNullOrEmpty(filePath) || string.IsNullOrEmpty(startCell))
            {
                MessageBox.Show("请提供完整的保存信息");
            }
            else
            {
                var t = Task.Run(() =>
                {
                    columnConfig = dtresult.AsEnumerable().GroupBy(n => new { name = n[columnName].ToString() }).Select(n => n.Key.name).ToList();
                    columnConfig.ForEach(f =>
                    {
                        DataTable dtfind = dtresult.Copy();
                        DataView dv = dtfind.DefaultView;
                        dv.RowFilter = $"{columnName}='{f}'";
                        var dtExport = dv.ToTable();
                        var path = $"{filePath}/{f}{fileName}.xlsx";
                        EppHelper.ExportByDt(path, dtExport, sheetName, flag, startCell);
                    });
                });
                Task.Run(() =>
                {
                    Task.WaitAll(t);
                    MessageBox.Show("导出完毕");
                });
            }
        }
    }
}
