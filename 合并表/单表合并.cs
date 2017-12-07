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
            dialog.Multiselect = true;//该值确定是否可以选择多个文件
            dialog.Title = "请选择文件夹";
            dialog.Filter = "所有文件(*.xls*)|*.xls*";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                CurrentPath = dialog.FileName;
                mcSheetNames.CbSource = null;
                try
                {
                    if (Path.GetExtension(CurrentPath) == "xlsx")
                    {
                        mcSheetNames.CbSource = EppHelper.GetAllSheet(CurrentPath);
                    }
                    else
                    {
                        mcSheetNames.CbSource = NpoiHelper.GetAllSheet(CurrentPath);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }

        private void btnMerge_Click(object sender, EventArgs e)
        {
            var sheets = mcSheetNames.TheValue;
            var sheetnames = string.IsNullOrEmpty(sheets) ? mcSheetNames.CbSource : sheets.Split(',').ToList();
            try
            {
                var dt = new DataTable();
                var ds = new DataSet();
                if (Path.GetExtension(CurrentPath) == "xlsx")
                {
                    ds = EppHelper.ReadExcelToDataSet(sheetnames, CurrentPath);
                }
                else
                {
                    ds = NpoiHelper.ReadExcelToDataSet(sheetnames, CurrentPath);
                }
                foreach (DataTable sdt in ds.Tables)
                {
                    dt.Merge(sdt);
                }
                dataGridView1.DataSource = dt;
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
                    var dt = dataGridView1.DataSource as DataTable;
                    if (EppHelper.ExportByDt(kk.FileName, dt))
                    {
                        MessageBox.Show("保存Excel成功");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }
    }
}
