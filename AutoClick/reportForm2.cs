using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;
namespace AutoClick
{
    public partial class reportForm2 : Form
    {
        public reportForm2()
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

        public void loadPOBalance()
        {
            this.dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            this.ContextMenuStrip = contextMenuStrip1;
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dt = pro.report_CustomerPOBalanceByType();
            dataGridView1.DataSource = dt;
            setRowNumber(dataGridView1);
            formatWeeklyPOBalanceByType(dataGridView1);

            dataGridView1.Columns["TOTAL_PO_BALANCE"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["TOTAL_PO_BALANCE"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["TOTAL_PO_BALANCE"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);
            if (!System.Windows.Forms.SystemInformation.TerminalServerSession)
            {
                Type dgvType = dataGridView1.GetType();
                PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                  BindingFlags.Instance | BindingFlags.NonPublic);
                pi.SetValue(dataGridView1, true, null);
            }
        }

        private void reportForm2_Load(object sender, EventArgs e)
        {

            loadPOBalance();

        }

        public void formatWeeklyPOBalanceByType(DataGridView dataGridView1)
        {
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;

            dataGridView1.Columns["TOTAL_PO_BALANCE"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["TSP"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["LABEL"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["UV"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["OLED"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["TAPE"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["RIBBON"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["SPT"].DefaultCellStyle.Format = "#,0";
        }

        private void saveTableToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new Form1().CopyToClipboardWithHeaders(dataGridView1);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            loadPOBalance();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelFactory.writeToExcelFile(dataGridView1);
        }
    }
}
