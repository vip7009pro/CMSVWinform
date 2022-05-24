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
    public partial class MaterialInfo : Form
    {
        public MaterialInfo()
        {
            InitializeComponent();
        }

        public string EMPL_NO;
        private void MaterialInfo_Load(object sender, EventArgs e)
        {
            this.ContextMenuStrip = contextMenuStrip1;
            if (!System.Windows.Forms.SystemInformation.TerminalServerSession)
            {
                Type dgvType = dataGridView1.GetType();
                PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                  BindingFlags.Instance | BindingFlags.NonPublic);
                pi.SetValue(dataGridView1, true, null);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            pro.insertMaterialfromBOMtoMTable();            
            dt = pro.getMaterialInfo(textBox1.Text, (checkBox1.Checked == true ? "thieuinfo" : "ok"));
            dataGridView1.DataSource = dt;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                ProductBLL pro = new ProductBLL();

                var selectedRows = dataGridView1.SelectedRows
                .OfType<DataGridViewRow>()
                .Where(row => !row.IsNewRow)
                .ToArray();

                foreach (var row in selectedRows)
                {
                    string
                        M_ID = row.Cells["M_ID"].Value.ToString(),
                        CUST_CD = row.Cells["CUST_CD"].Value.ToString(),
                        SSPRICE = row.Cells["SSPRICE"].Value.ToString(),
                        CMSPRICE = row.Cells["CMSPRICE"].Value.ToString(),
                        SLITTING_PRICE = row.Cells["SLITTING_PRICE"].Value.ToString(),
                        MASTER_WIDTH = row.Cells["MASTER_WIDTH"].Value.ToString(),
                        ROLL_LENGTH = row.Cells["ROLL_LENGTH"].Value.ToString();
                    string updateValue = $" SET CUST_CD='{CUST_CD}', SSPRICE='{SSPRICE}', CMSPRICE='{CMSPRICE}', SLITTING_PRICE='{SLITTING_PRICE}',MASTER_WIDTH='{MASTER_WIDTH}',ROLL_LENGTH='{ROLL_LENGTH}' WHERE M_ID={M_ID}";
                    pro.updateMaterial(updateValue);
                }
                MessageBox.Show("Update Material info thành công !");

            }
            catch(Exception ex)
            {
                MessageBox.Show("Lỗi update Material: " + ex.ToString());
            }
           


        }

        private void updateToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
        public bool sort_on_off = true;
        private void button3_Click(object sender, EventArgs e)
        {
            if(sort_on_off)
            {
                foreach (DataGridViewColumn column in dataGridView1.Columns)
                {
                    column.SortMode = DataGridViewColumnSortMode.NotSortable;
                }
                sort_on_off = false;

            }
            else
            {
                foreach (DataGridViewColumn column in dataGridView1.Columns)
                {
                    column.SortMode = DataGridViewColumnSortMode.Automatic;
                }
                sort_on_off = true;
            }
            
        }
    }
}
