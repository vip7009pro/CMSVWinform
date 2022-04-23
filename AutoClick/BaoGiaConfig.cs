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
    public partial class BaoGiaConfig : Form
    {
        public BaoGiaConfig()
        {
            InitializeComponent();
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        public void searchcodegia()
        {
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            pro.insertMaterialfromBOMtoMTable();
            dt = pro.getBaoGiaConfig();
            dataGridView1.DataSource = dt;           

        }

        private void button1_Click(object sender, EventArgs e)
        {
            searchcodegia();
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
                        CONFIG_ID = row.Cells["CONFIG_ID"].Value.ToString(),
                        PROD_TYPE = row.Cells["PROD_TYPE"].Value.ToString(),
                        EQ = row.Cells["EQ"].Value.ToString(),
                        SIZE = row.Cells["SIZE"].Value.ToString(),
                        STEP = row.Cells["STEP"].Value.ToString(),
                        INK_COST = row.Cells["INK_COST"].Value.ToString(),
                        INSPECTION_COST = row.Cells["INSPECTION_COST"].Value.ToString(),
                        NG_RATE = row.Cells["NG_RATE"].Value.ToString(),
                        LABOR_DEPRE_COST = row.Cells["LABOR_DEPRE_COST"].Value.ToString();

                    string updatevalue = $" SET PROD_TYPE='{PROD_TYPE}', EQ='{EQ}', SIZE='{SIZE}', STEP='{SIZE}', INK_COST='{INK_COST}', INSPECTION_COST='{INSPECTION_COST}', NG_RATE='{NG_RATE}', LABOR_DEPRE_COST='{LABOR_DEPRE_COST}' WHERE CONFIG_ID={CONFIG_ID}";

                   // pro.updateConfig(updatevalue);

                }
                dataGridView1.ClearSelection();
                MessageBox.Show("Update config giá thành công !");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi update config giá: " + ex.ToString());
            }
        }

        private void BaoGiaConfig_Load(object sender, EventArgs e)
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
