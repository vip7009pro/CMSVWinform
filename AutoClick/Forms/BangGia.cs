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
    public partial class BangGia : Form
    {
        public BangGia()
        {
            InitializeComponent();
        }

        public void searchcodegia()
        {
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            pro.insertMaterialfromBOMtoMTable();
            dt = pro.getcodebom2Info(textBox1.Text, (checkBox1.Checked == true ? "chuatinhgia" : "tinhgiaroi"));
            dataGridView1.DataSource = dt;
            formatcodelist(dataGridView1);

        }
        public void searchfullBOMgia()
        {
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            pro.insertMaterialfromBOMtoMTable();
            dt = pro.getgiafullBOM2Info(textBox1.Text,"lieu");
            dataGridView1.DataSource = dt;
            if(dt.Rows.Count > 0)
            {
                dataGridView1.Rows[0].DefaultCellStyle.BackColor = Color.LightGray;
                int lastcolor = 1;

                for (int i = 1; i < dt.Rows.Count; i++)
                {
                    DataGridViewRow lastrow = dataGridView1.Rows[i - 1];
                    DataGridViewRow thisrow = dataGridView1.Rows[i];
                    string lastname = lastrow.Cells["G_CODE"].Value.ToString();
                    string thisname = thisrow.Cells["G_CODE"].Value.ToString();

                    if (thisname != lastname)
                    {
                        if (lastcolor == 1)
                        {
                            lastcolor = 0;
                            thisrow.DefaultCellStyle.BackColor = Color.White;
                        }
                        else
                        {
                            lastcolor = 1;
                            thisrow.DefaultCellStyle.BackColor = Color.LightGray;
                        }

                    }
                    else
                    {
                        thisrow.DefaultCellStyle.BackColor = lastrow.DefaultCellStyle.BackColor;
                    }


                }
                this.dataGridView1.DefaultCellStyle.BackColor = Color.Beige;
                this.dataGridView1.DefaultCellStyle.Font = new Font("Tahoma", 12);

                //formatcodelist(dataGridView1);
                dataGridView1.Columns["G_NAME"].DefaultCellStyle.BackColor = Color.Brown;
                dataGridView1.Columns["G_NAME"].DefaultCellStyle.ForeColor = Color.Yellow;

            }
            
            
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if(checkBox2.Checked == true)
            {
                searchfullBOMgia();
            }
            else
            {
                searchcodegia();
            }
            
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (checkBox2.Checked == true)
                {
                    searchfullBOMgia();
                }
                else
                {
                    searchcodegia();
                }
            }
        }

        public void formatcodelist(DataGridView dataGridView1)
        {
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;

            dataGridView1.Columns["CUST_NAME_KD"].DefaultCellStyle.BackColor = Color.Aqua;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.BackColor = Color.Brown;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.ForeColor = Color.Yellow;


            dataGridView1.Columns["MATERIAL_COST_CMS"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["MATERIAL_COST_CMS"].DefaultCellStyle.ForeColor = Color.White;

            dataGridView1.Columns["PROCESS_COST_CMS"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["PROCESS_COST_CMS"].DefaultCellStyle.ForeColor = Color.White;

            dataGridView1.Columns["OTHER_COST_CMS"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["OTHER_COST_CMS"].DefaultCellStyle.ForeColor = Color.White;

            dataGridView1.Columns["PROFIT_VALUE_CMS"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["PROFIT_VALUE_CMS"].DefaultCellStyle.ForeColor = Color.White;

            dataGridView1.Columns["MCR_CMS"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["MCR_CMS"].DefaultCellStyle.ForeColor = Color.White;

            dataGridView1.Columns["MATERIAL_COST_SS"].DefaultCellStyle.BackColor = Color.DarkGray;
            dataGridView1.Columns["MATERIAL_COST_SS"].DefaultCellStyle.ForeColor = Color.White;

            dataGridView1.Columns["PROCESS_COST_SS"].DefaultCellStyle.BackColor = Color.DarkGray;
            dataGridView1.Columns["PROCESS_COST_SS"].DefaultCellStyle.ForeColor = Color.White;

            dataGridView1.Columns["OTHER_COST_SS"].DefaultCellStyle.BackColor = Color.DarkGray;
            dataGridView1.Columns["OTHER_COST_SS"].DefaultCellStyle.ForeColor = Color.White;

            dataGridView1.Columns["PROFIT_VALUE_SS"].DefaultCellStyle.BackColor = Color.DarkGray;
            dataGridView1.Columns["PROFIT_VALUE_SS"].DefaultCellStyle.ForeColor = Color.White;

            dataGridView1.Columns["MCR_SS"].DefaultCellStyle.BackColor = Color.DarkGray;
            dataGridView1.Columns["MCR_SS"].DefaultCellStyle.ForeColor = Color.White;

            dataGridView1.Columns["PRODUCT_CMSPRICE"].DefaultCellStyle.BackColor = Color.Orange;
            dataGridView1.Columns["PRODUCT_CMSPRICE"].DefaultCellStyle.ForeColor = Color.White;

            dataGridView1.Columns["PRODUCT_SSPRICE"].DefaultCellStyle.BackColor = Color.LightBlue;
            dataGridView1.Columns["PRODUCT_SSPRICE"].DefaultCellStyle.ForeColor = Color.Black;

            dataGridView1.Columns["BEP_PRICE"].DefaultCellStyle.BackColor = Color.Orange;
            dataGridView1.Columns["BEP_PRICE"].DefaultCellStyle.ForeColor = Color.White;
        }

        private void BangGia_Load(object sender, EventArgs e)
        {
            if (!System.Windows.Forms.SystemInformation.TerminalServerSession)
            {
                Type dgvType = dataGridView1.GetType();
                PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                  BindingFlags.Instance | BindingFlags.NonPublic);
                pi.SetValue(dataGridView1, true, null);
            }
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
            this.dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
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
                        PRODUCT_FINAL_PRICE = row.Cells["PRODUCT_FINAL_PRICE"].Value.ToString();

                    string updatevalue = $" SET PRODUCT_FINAL_PRICE='{PRODUCT_FINAL_PRICE}' WHERE G_CODE={G_CODE}";

                    pro.updatebaogiaM100(updatevalue);

                }
                dataGridView1.ClearSelection();
                MessageBox.Show("Update giá chốt thành công !");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi update giá chốt: " + ex.ToString());
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelFactory.writeToExcelFile(dataGridView1);
        }
    }
}
