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
    public partial class BEP_Info : Form
    {
        public BEP_Info()
        {
            InitializeComponent();
        }
        public void loadbeplist()
        {
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            pro.insertMaterialfromBOMtoMTable();
            dt = pro.getbeplist(textBox1.Text, checkBox2.Checked == true ? "chuabep" : "other");
            dataGridView1.DataSource = dt;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            loadbeplist();
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            loadbeplist();
        }

        private void BEP_Info_Load(object sender, EventArgs e)
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

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();

                var selectedRows = dataGridView1.SelectedRows
                .OfType<DataGridViewRow>()
                .Where(row => !row.IsNewRow)
                .ToArray();

                foreach (var row in selectedRows)
                {
                    string
                        G_CODE = row.Cells["G_CODE"].Value.ToString(),
                        PROD_MANPOWER = row.Cells["PROD_MANPOWER"].Value.ToString(),
                        INSPECT_MANPOWER = row.Cells["INSPECT_MANPOWER"].Value.ToString(),
                        BEP_1HOUR_PROD_QTY = row.Cells["BEP_1HOUR_PROD_QTY"].Value.ToString(),
                        BEP_PROD_NG_RATE = row.Cells["BEP_PROD_NG_RATE"].Value.ToString(),
                        BEP_INSP_NG_RATE = row.Cells["BEP_INSP_NG_RATE"].Value.ToString();                       

                    string updatevalue = $" SET PROD_MANPOWER='{PROD_MANPOWER}', BEP_1HOUR_PROD_QTY='{BEP_1HOUR_PROD_QTY}', BEP_PROD_NG_RATE='{BEP_PROD_NG_RATE}' WHERE G_CODE='{G_CODE}'";

                     pro.updateBEPInfo(updatevalue);

                }
                dataGridView1.ClearSelection();
                MessageBox.Show("Update BEP info thành công !");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi update BEP Info: " + ex.ToString());
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();

                var selectedRows = dataGridView1.SelectedRows
                .OfType<DataGridViewRow>()
                .Where(row => !row.IsNewRow)
                .ToArray();

                foreach (var row in selectedRows)
                {
                    string
                        G_CODE = row.Cells["G_CODE"].Value.ToString(),
                        PROD_MANPOWER = row.Cells["PROD_MANPOWER"].Value.ToString(),
                        INSPECT_MANPOWER = row.Cells["INSPECT_MANPOWER"].Value.ToString(),
                        BEP_1HOUR_PROD_QTY = row.Cells["BEP_1HOUR_PROD_QTY"].Value.ToString(),
                        BEP_PROD_NG_RATE = row.Cells["BEP_PROD_NG_RATE"].Value.ToString(),
                        BEP_INSP_NG_RATE = row.Cells["BEP_INSP_NG_RATE"].Value.ToString();

                    string updatevalue = $" SET INSPECT_MANPOWER='{INSPECT_MANPOWER}', BEP_INSP_NG_RATE='{BEP_INSP_NG_RATE}' WHERE G_CODE='{G_CODE}'";
                    pro.updateBEPInfo(updatevalue);

                }
                dataGridView1.ClearSelection();
                MessageBox.Show("Update BEP info thành công !");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi update BEP Info: " + ex.ToString());
            }

        }
    }
}
