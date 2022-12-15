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
    public partial class SOPFULL : Form
    {
        public SOPFULL()
        {
            InitializeComponent();
        }

        public int tra_gcode_flag = 0;
        public int tra_g_name_kd_flag = 0;

        public DataTable dtgv1_data = new DataTable();
       

        public string generate_condition()
        {
            string cd = " WHERE 1=1 ";
            string G_NAME = "";
            if(textBox1.Text != "")
            {
                G_NAME = $" AND TONKHOFULL.G_NAME LIKE '%{textBox1.Text}%'";
            }
            string G_CODE = "";
            if (textBox2.Text != "")
            {
                G_CODE = $" AND PO_TABLE_1.G_CODE = '{textBox2.Text}'";
            }
            string THUA_THIEU = "";
            if(checkBox1.Checked== true)
            {
                THUA_THIEU = " AND PO_TABLE_1.PO_BALANCE >0 ";
            }
            cd += G_CODE + G_NAME + THUA_THIEU;
            return cd;
        }

        public void formatpotonkho(DataGridView dataGridView1)
        {
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;

            dataGridView1.Columns["PO_QTY"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["TOTAL_DELIVERED"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["PO_BALANCE"].DefaultCellStyle.Format = "#,0";
           

            dataGridView1.Columns["TONG_TON_KIEM"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["BTP"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["TON_TP"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["BLOCK_QTY"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["GRAND_TOTAL_STOCK"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["CHO_KIEM"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["CHO_CS_CHECK"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["CHO_KIEM_RMA"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["THUA_THIEU"].DefaultCellStyle.Format = "#,0";



            dataGridView1.Columns["PO_QTY"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["PO_QTY"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["PO_QTY"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["PO_BALANCE"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["PO_BALANCE"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["PO_BALANCE"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["TOTAL_DELIVERED"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["TOTAL_DELIVERED"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["TOTAL_DELIVERED"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

           

            
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.BackColor = Color.Brown;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.ForeColor = Color.Yellow;

            dataGridView1.Columns["GRAND_TOTAL_STOCK"].DefaultCellStyle.ForeColor = Color.Black;
            dataGridView1.Columns["GRAND_TOTAL_STOCK"].DefaultCellStyle.BackColor = Color.GreenYellow;
            dataGridView1.Columns["GRAND_TOTAL_STOCK"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);


            dataGridView1.Columns["TONG_TON_KIEM"].DefaultCellStyle.ForeColor = Color.Blue;
            dataGridView1.Columns["TONG_TON_KIEM"].DefaultCellStyle.BackColor = Color.Aqua;
            dataGridView1.Columns["TONG_TON_KIEM"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Regular);

            dataGridView1.Columns["BTP"].DefaultCellStyle.ForeColor = Color.Blue;
            dataGridView1.Columns["BTP"].DefaultCellStyle.BackColor = Color.Aqua;
            dataGridView1.Columns["BTP"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Regular);

            dataGridView1.Columns["TON_TP"].DefaultCellStyle.ForeColor = Color.Blue;
            dataGridView1.Columns["TON_TP"].DefaultCellStyle.BackColor = Color.Aqua;
            dataGridView1.Columns["TON_TP"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Regular);

            dataGridView1.Columns["BLOCK_QTY"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["BLOCK_QTY"].DefaultCellStyle.BackColor = Color.Red;
            dataGridView1.Columns["BLOCK_QTY"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Regular);

            dataGridView1.Columns["THUA_THIEU"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["THUA_THIEU"].DefaultCellStyle.BackColor = Color.DarkRed;
            dataGridView1.Columns["THUA_THIEU"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Regular);



        }

        public void formatpotonkhoKD(DataGridView dataGridView1)
        {
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;

            dataGridView1.Columns["PO_QTY"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["TOTAL_DELIVERED"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["PO_BALANCE"].DefaultCellStyle.Format = "#,0";


            dataGridView1.Columns["TONG_TON_KIEM"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["BTP"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["TON_TP"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["BLOCK_QTY"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["GRAND_TOTAL_STOCK"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["CHO_KIEM"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["CHO_CS_CHECK"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["CHO_KIEM_RMA"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["THUA_THIEU"].DefaultCellStyle.Format = "#,0";



            dataGridView1.Columns["PO_QTY"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["PO_QTY"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["PO_QTY"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["PO_BALANCE"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["PO_BALANCE"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["PO_BALANCE"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["TOTAL_DELIVERED"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["TOTAL_DELIVERED"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["TOTAL_DELIVERED"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);



            dataGridView1.Columns["G_NAME_KD"].DefaultCellStyle.BackColor = Color.Brown;
            dataGridView1.Columns["G_NAME_KD"].DefaultCellStyle.ForeColor = Color.Yellow;

            dataGridView1.Columns["GRAND_TOTAL_STOCK"].DefaultCellStyle.ForeColor = Color.Black;
            dataGridView1.Columns["GRAND_TOTAL_STOCK"].DefaultCellStyle.BackColor = Color.GreenYellow;
            dataGridView1.Columns["GRAND_TOTAL_STOCK"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);


            dataGridView1.Columns["TONG_TON_KIEM"].DefaultCellStyle.ForeColor = Color.Blue;
            dataGridView1.Columns["TONG_TON_KIEM"].DefaultCellStyle.BackColor = Color.Aqua;
            dataGridView1.Columns["TONG_TON_KIEM"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Regular);

            dataGridView1.Columns["BTP"].DefaultCellStyle.ForeColor = Color.Blue;
            dataGridView1.Columns["BTP"].DefaultCellStyle.BackColor = Color.Aqua;
            dataGridView1.Columns["BTP"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Regular);

            dataGridView1.Columns["TON_TP"].DefaultCellStyle.ForeColor = Color.Blue;
            dataGridView1.Columns["TON_TP"].DefaultCellStyle.BackColor = Color.Aqua;
            dataGridView1.Columns["TON_TP"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Regular);

            dataGridView1.Columns["BLOCK_QTY"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["BLOCK_QTY"].DefaultCellStyle.BackColor = Color.Red;
            dataGridView1.Columns["BLOCK_QTY"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Regular);

            dataGridView1.Columns["THUA_THIEU"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["THUA_THIEU"].DefaultCellStyle.BackColor = Color.DarkRed;
            dataGridView1.Columns["THUA_THIEU"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Regular);



        }
        private void button3_Click(object sender, EventArgs e)
        {
            tra_gcode_flag = 1;
            tra_g_name_kd_flag = 0;
            if (!backgroundWorker1.IsBusy)
            {
                pictureBox1.Show();
                backgroundWorker1.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Đang trong tiến trình khác, thử lại sau");
            }

        }

        public void traPO_TONKHO()
        {
            ProductBLL pro = new ProductBLL();
            if(tra_gcode_flag == 1)
            {               
                dtgv1_data = pro.traPO_TONKHO(generate_condition());
            }
            else if(tra_g_name_kd_flag == 1)
            {               
                dtgv1_data = pro.traPO_TONKHOKD(generate_condition()); 
            }

        }


        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {            
                traPO_TONKHO();           
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {

        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        { 
            if (tra_gcode_flag == 1)
            {
                dataGridView1.Columns.Clear();
                dataGridView1.DataSource = dtgv1_data;
                formatpotonkho(dataGridView1);
                new Form1().setRowNumber(dataGridView1);
                pictureBox1.Hide();
                MessageBox.Show("Đã load: " + dtgv1_data.Rows.Count + " dòng");

            }
            else if(tra_g_name_kd_flag == 1)
            {
                dataGridView1.Columns.Clear();
                dataGridView1.DataSource = dtgv1_data;
                formatpotonkhoKD(dataGridView1);
                new Form1().setRowNumber(dataGridView1);
                pictureBox1.Hide();
                MessageBox.Show("Đã load: " + dtgv1_data.Rows.Count + " dòng");
            }

            ProductBLL pro = new ProductBLL();
            dtgv1_data = pro.report_checkUP_BTP_today();
            if (dtgv1_data.Rows.Count > 0)
            {
                label8.Text = "BTP: Đã UP";
            }
            else
            {
                label8.Text = "BTP: Chưa UP";
            }

            dtgv1_data = pro.report_checkUP_TONKIEM_today();
            if (dtgv1_data.Rows.Count > 0)
            {
                label7.Text = "Tồn kiểm: Đã UP";
            }
            else
            {
                label7.Text = "Tồn kiểm: Chưa UP";
            }

        }

        private void SOPFULL_Load(object sender, EventArgs e)
        {
            if (!System.Windows.Forms.SystemInformation.TerminalServerSession)
            {
                Type dgvType = dataGridView1.GetType();
                PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                  BindingFlags.Instance | BindingFlags.NonPublic);
                pi.SetValue(dataGridView1, true, null);
            }
            pictureBox1.Hide();
            this.dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
        }

        private void button1_Click(object sender, EventArgs e)
        {           
            tra_gcode_flag = 0;
            tra_g_name_kd_flag = 1;
            if (!backgroundWorker1.IsBusy)
            {
                pictureBox1.Show();
                backgroundWorker1.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Đang trong tiến trình khác, thử lại sau");
            }
        }
    }
}
