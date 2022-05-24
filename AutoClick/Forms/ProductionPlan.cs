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
    public partial class ProductionPlan : Form
    {
        public ProductionPlan()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dt = pro.getProductionInfo();
            dataGridView1.DataSource = dt;
            formatdtgv(dataGridView1);

        }
        public void formatdtgv(DataGridView dataGridView1)
        {
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;

            dataGridView1.Columns["CUST_NAME_KD"].DefaultCellStyle.BackColor = Color.Aqua;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.BackColor = Color.Green;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.ForeColor = Color.Yellow;

            dataGridView1.Columns["G_NAME_KD"].DefaultCellStyle.BackColor = Color.Brown;
            dataGridView1.Columns["G_NAME_KD"].DefaultCellStyle.ForeColor = Color.Yellow;

            dataGridView1.Columns["LAST_ACTIVE_TIME"].DefaultCellStyle.BackColor = Color.Pink;
            dataGridView1.Columns["LAST_ACTIVE_TIME"].DefaultCellStyle.ForeColor = Color.Blue;


        }

        private void ProductionPlan_Load(object sender, EventArgs e)
        {
            if (!System.Windows.Forms.SystemInformation.TerminalServerSession)
            {
                Type dgvType = dataGridView1.GetType();
                PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                  BindingFlags.Instance | BindingFlags.NonPublic);
                pi.SetValue(dataGridView1, true, null);
            }
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
        }
    }
}
