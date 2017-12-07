using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PopupTool;

namespace 合并表
{
    public partial class MultipleCombobox : UserControl
    {
        public MultipleCombobox()
        {
            InitializeComponent();
            this.Size = new Size(220, 38);
        }
        public TextBox TextBox
        {
            get { return this.textBox1; }
        }
        public string ColumnName { get; set; }
        private string _title;
        public string Title
        {
            get { return _title; }
            set
            {
                _title = value;
                if (_title != null)
                {
                    this.label1.Text = _title + "↓：";
                    this.textBox1.Location = new Point(label1.Width + 9, label1.Location.Y-3);
                    this.textBox1.Size = new Size(this.Width - label1.Width - 12, 38);
                }
            }
        }

        private List<string> _cbSource;
        public List<string> CbSource
        {
            get { return _cbSource; }
            set
            {
                _cbSource = value;
            }
        }

        public string TheValue
        {
            get { return textBox1.Text.Trim().Replace('，', ','); }
        }

        TreeView tv;

        private void label1_Click(object sender, EventArgs e)
        {
            tv = new TreeView();
            tv.CheckBoxes = true;
            if (_cbSource != null)
            {
                tv.Nodes.Clear();
                _cbSource.ForEach(n => tv.Nodes.Add(n, n));
                tv.AfterCheck += tv_AfterCheck;
                tv.NodeMouseDoubleClick += tv_NodeMouseDoubleClick;
                if (!string.IsNullOrEmpty(TheValue))
                {
                    var s = TheValue.Split(',');
                    s.ToList().ForEach(n =>
                    {
                        if (tv.Nodes.Find(n, true).Count() > 0)
                            tv.Nodes.Find(n, true)[0].Checked = true;
                    });
                }
            }
            var pp = new Popup(tv);
            pp.Size = new Size(220, 180);
            pp.BackColor = System.Drawing.Color.Transparent;
            tv.Dock = DockStyle.Fill;
            pp.Show(this);
        }

        void tv_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            e.Node.Checked = !e.Node.Checked;
        }

        private void tv_AfterCheck(object sender, TreeViewEventArgs e)
        {
            string result = "";
            foreach (TreeNode tn in tv.Nodes)
            {
                if (tn.Checked)
                {
                    result += string.IsNullOrEmpty(result) ? tn.Text : String.Format(",{0}", tn.Text);
                }
            }
            textBox1.Text = result;
        }
    }
}
