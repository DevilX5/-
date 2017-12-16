using OfficeOpenXml;
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
    public partial class 单表合并 : Form
    {
        public 单表合并()
        {
            InitializeComponent();
        }
        public string CurrentPath { get; set; }
        private void btnChoose_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;
            dialog.Title = "请选择文件夹";
            dialog.Filter = "所有文件(*.xls*)|*.xls*";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                CurrentPath = dialog.FileName;
                mcSheetNames.CbSource = null;
                label1.Text = "正在获取sheet列表";
                var t = Task.Run(() =>
                {
                    return ExcelHelper.GetSheetList(CurrentPath);
                });
                Task.Run(() => 
                {
                    Task.WaitAll(t);
                    this.Invoke((MethodInvoker)delegate
                    {
                        mcSheetNames.CbSource = t.Result;
                        label1.Text = "sheet列表获取完毕";
                    });
                });
            }
        }

        private void btnMerge_Click(object sender, EventArgs e)
        {
            var fs = new FileAndSheets();
            fs.FileName = CurrentPath;
            var sheets = mcSheetNames.TheValue;
            fs.SheetNames = string.IsNullOrEmpty(sheets) ? mcSheetNames.CbSource : sheets.Split(',').ToList();
            try
            {
                var dt = new DataTable();
                //var ds = new DataSet();
                label1.Text = "正在合并sheet内容";
                var t = Task.Run(() =>
                {
                    var ds = new DataSet("ds");
                    ExcelHelper.SetDataSet(ds,fs);
                    return ds;
                });
                Task.Run(() =>
                {
                    Task.WaitAll(t);
                    foreach (DataTable sdt in t.Result.Tables)
                    {
                        dt.Merge(sdt);
                    }
                    this.Invoke((MethodInvoker)delegate
                    {
                        dataGridView1.DataSource = dt;
                        label1.Text = "数据合并完毕";
                    });
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            SaveFileDialog kk = new SaveFileDialog();
            kk.Title = "保存EXCEL文件";
            kk.Filter = "excel文件(*.xlsx) |*.xlsx ";
            kk.FilterIndex = 1;
            if (kk.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    label1.Text = "正在导出数据";
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
                                label1.Text = "导出成功";
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

        private void dataGridView1_RowStateChanged(object sender, DataGridViewRowStateChangedEventArgs e)
        {
            e.Row.HeaderCell.Value = string.Format("{0}", e.Row.Index + 1);
        }
    }
}
