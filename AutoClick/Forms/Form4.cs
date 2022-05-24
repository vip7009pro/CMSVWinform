using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoClick
{
    public partial class Form4 : Form
    {
        public Form4()
        {
            InitializeComponent();
        }

        private void setRowNumber(DataGridView dgv)
        {
            foreach (DataGridViewRow row in dgv.Rows)
            {
                row.HeaderCell.Value = (row.Index + 1).ToString();
            }
        }

        public void initdata()
        {
            ProductBLL pro = new ProductBLL();
            DataTable dtoverduepic = new DataTable();
            DataTable dtoverduecustomer = new DataTable();
            DataTable dtoverduetype = new DataTable();

            dtoverduepic = pro.report_OverDueByPIC();
            dtoverduecustomer = pro.report_OverDueByCustomer();
            dtoverduetype = pro.report_OverDueByTYPE();

            dataGridView1.DataSource = dtoverduepic;
            dataGridView2.DataSource = dtoverduetype;
            dataGridView3.DataSource = dtoverduecustomer;
            setRowNumber(dataGridView1);
            setRowNumber(dataGridView2);
            setRowNumber(dataGridView3);



        }

        private void Form4_Load(object sender, EventArgs e)
        {
            initdata();
            this.ContextMenuStrip = contextMenuStrip1;
            this.dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            this.dataGridView2.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            this.dataGridView3.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
        }

        private void save1ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new Form1().CopyToClipboardWithHeaders(dataGridView1);
        }

        private void save2ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new Form1().CopyToClipboardWithHeaders(dataGridView2);
        }

        private void save3ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new Form1().CopyToClipboardWithHeaders(dataGridView3);
        }
    }
}
