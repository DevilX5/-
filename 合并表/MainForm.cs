using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 合并表
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var f = new 单表合并();
            f.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var f = new 合并文件夹();
            f.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var f = new 分表();
            f.Show();
        }
    }
}
