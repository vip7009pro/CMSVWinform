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
    public partial class CheckBOMGia : Form
    {
        public CheckBOMGia()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            pro.insertMaterialfromBOMtoMTable();
            dt = pro.getBOM2Info(textBox1.Text, checkBox1.Checked==true ? "lieu":"other");
            dataGridView1.DataSource = dt;
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
                    if(lastcolor == 1)
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


        }

        private void CheckBOMGia_Load(object sender, EventArgs e)
        {
            if (!System.Windows.Forms.SystemInformation.TerminalServerSession)
            {
                Type dgvType = dataGridView1.GetType();
                PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                  BindingFlags.Instance | BindingFlags.NonPublic);
                pi.SetValue(dataGridView1, true, null);
            }

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
                        BOM_ID = row.Cells["BOM_ID"].Value.ToString(),
                        M_NAME = row.Cells["M_NAME"].Value.ToString();
                        dt = pro.getMaterialInfo(M_NAME, "ok");
                    if (dt.Rows.Count > 0)
                    {
                        row.Cells["CUST_CD"].Value = dt.Rows[0]["CUST_CD"];
                        row.Cells["M_CMS_PRICE"].Value = dt.Rows[0]["CMSPRICE"];
                        row.Cells["M_SS_PRICE"].Value = dt.Rows[0]["SSPRICE"];
                        row.Cells["M_SLITTING_PRICE"].Value = dt.Rows[0]["SLITTING_PRICE"];
                        row.Cells["MAT_MASTER_WIDTH"].Value = dt.Rows[0]["MASTER_WIDTH"];
                        row.Cells["MAT_ROLL_LENGTH"].Value = dt.Rows[0]["ROLL_LENGTH"];

                        row.Cells["CUST_CD"].Style.BackColor = Color.LightGray;
                        row.Cells["M_CMS_PRICE"].Style.BackColor = Color.LightGray;
                        row.Cells["M_SS_PRICE"].Style.BackColor = Color.LightGray;
                        row.Cells["M_SLITTING_PRICE"].Style.BackColor = Color.LightGray;
                        row.Cells["MAT_MASTER_WIDTH"].Style.BackColor = Color.LightGray;
                        row.Cells["MAT_ROLL_LENGTH"].Style.BackColor = Color.LightGray;

                    }
                   
                    //pro.updateMaterial(updateValue);
                }
                dataGridView1.ClearSelection();
                MessageBox.Show("Cập nhật bom giá hoàn thành !");

            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi update Material: " + ex.ToString());
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
                        BOM_ID = row.Cells["BOM_ID"].Value.ToString(),
                        CUST_CD = row.Cells["CUST_CD"].Value.ToString(),
                        M_CMS_PRICE = row.Cells["M_CMS_PRICE"].Value.ToString(),
                        M_SS_PRICE = row.Cells["M_SS_PRICE"].Value.ToString(),
                        M_SLITTING_PRICE = row.Cells["M_SLITTING_PRICE"].Value.ToString(),
                        MAT_MASTER_WIDTH = row.Cells["MAT_MASTER_WIDTH"].Value.ToString(),
                        MAT_ROLL_LENGTH = row.Cells["MAT_ROLL_LENGTH"].Value.ToString(),
                        USAGE = row.Cells["USAGE"].Value.ToString(),
                        MAT_CUTWIDTH = row.Cells["MAT_CUTWIDTH"].Value.ToString();


                    string updatevalue = $" SET CUST_CD='{CUST_CD}', M_CMS_PRICE='{M_CMS_PRICE}', M_SS_PRICE='{M_SS_PRICE}', M_SLITTING_PRICE='{M_SLITTING_PRICE}', MAT_MASTER_WIDTH='{MAT_MASTER_WIDTH}', MAT_ROLL_LENGTH='{MAT_ROLL_LENGTH}', USAGE='{USAGE}', MAT_CUTWIDTH='{MAT_CUTWIDTH}' WHERE BOM_ID={BOM_ID}";
                   
                    pro.updateBOM2(updatevalue);

                }
                dataGridView1.ClearSelection();
                MessageBox.Show("Update Material info thành công !");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi update Material: " + ex.ToString());
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (dataGridView1 == null || dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("Tra bom trước");
            }
            else
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (!row.IsNewRow)
                    {
                        if (row.Cells["CUST_CD"].Value.ToString() == "") row.Cells["CUST_CD"].Style.BackColor = Color.Red;
                        if (row.Cells["IMPORT_CAT"].Value.ToString() == "") row.Cells["IMPORT_CAT"].Style.BackColor = Color.Red;
                        if (row.Cells["M_CMS_PRICE"].Value.ToString() == "0") row.Cells["M_CMS_PRICE"].Style.BackColor = Color.Red;
                        if (row.Cells["M_SS_PRICE"].Value.ToString() == "0") row.Cells["M_SS_PRICE"].Style.BackColor = Color.Red;
                        if (row.Cells["USAGE"].Value.ToString() == "") row.Cells["USAGE"].Style.BackColor = Color.Red;
                        if (row.Cells["MAT_MASTER_WIDTH"].Value.ToString() == "0") row.Cells["MAT_MASTER_WIDTH"].Style.BackColor = Color.Red;
                        if (row.Cells["MAT_CUTWIDTH"].Value.ToString() == "0") row.Cells["MAT_CUTWIDTH"].Style.BackColor = Color.Red;
                        if (row.Cells["MAT_ROLL_LENGTH"].Value.ToString() == "0") row.Cells["MAT_ROLL_LENGTH"].Style.BackColor = Color.Red;
                        if (row.Cells["MAT_ROLL_LENGTH"].Value.ToString() == "0") row.Cells["MAT_ROLL_LENGTH"].Style.BackColor = Color.Red;
                    }
                }
            }
        }
    }
}
