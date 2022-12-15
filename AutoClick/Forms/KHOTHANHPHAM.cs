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
    public partial class KHOTHANHPHAM : Form
    {
        public KHOTHANHPHAM()
        {
            InitializeComponent();
        }

        public int input_flag = 0;
        public int output_flag = 0;
        public int ton_gop_flag = 0;
        public int ton_gop2_flag = 0;
        public int ton_tach_flag = 0;
        public DataTable dtgv1_data = new DataTable();
        
        public void formatDataGridViewtraSTOCK_IN(DataGridView dataGridView1)
        {
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;

            dataGridView1.Columns["IO_Qty"].DefaultCellStyle.Format = "#,0";
           

            dataGridView1.Columns["Product_MaVach"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["Product_MaVach"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["Product_MaVach"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

     
            dataGridView1.Columns["IO_Qty"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["IO_Qty"].DefaultCellStyle.BackColor = Color.Black;
            dataGridView1.Columns["IO_Qty"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["G_NAME"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.BackColor = Color.Black;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);           

            dataGridView1.Columns["Customer_ShortName"].DefaultCellStyle.BackColor = Color.Aqua;


        }

        public void formatDataGridViewtraSTOCK_OUT(DataGridView dataGridView1)
        {
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;

            dataGridView1.Columns["IO_Qty"].DefaultCellStyle.Format = "#,0";


            dataGridView1.Columns["Product_MaVach"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["Product_MaVach"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["Product_MaVach"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);


            dataGridView1.Columns["IO_Qty"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["IO_Qty"].DefaultCellStyle.BackColor = Color.Black;
            dataGridView1.Columns["IO_Qty"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["G_NAME"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.BackColor = Color.Black;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["Customer_ShortName"].DefaultCellStyle.BackColor = Color.Aqua;


        }


        public void formatDataGridViewtraTON_GOP(DataGridView dataGridView1)
        {
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;

            dataGridView1.Columns["CHO_KIEM"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["CHO_CS_CHECK"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["CHO_KIEM_RMA"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["TONG_TON_KIEM"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["BTP"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["TON_TP"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["BLOCK_QTY"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["GRAND_TOTAL_STOCK"].DefaultCellStyle.Format = "#,0";


            dataGridView1.Columns["G_CODE"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["G_CODE"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["G_CODE"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);


            dataGridView1.Columns["G_NAME"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.BackColor = Color.Black;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["GRAND_TOTAL_STOCK"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["GRAND_TOTAL_STOCK"].DefaultCellStyle.BackColor = Color.Black;
            dataGridView1.Columns["GRAND_TOTAL_STOCK"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);


            dataGridView1.Columns["TONG_TON_KIEM"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["TONG_TON_KIEM"].DefaultCellStyle.BackColor = Color.Blue;
            dataGridView1.Columns["TONG_TON_KIEM"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Regular);

            dataGridView1.Columns["BTP"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["BTP"].DefaultCellStyle.BackColor = Color.Blue;
            dataGridView1.Columns["BTP"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Regular);

            dataGridView1.Columns["TON_TP"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["TON_TP"].DefaultCellStyle.BackColor = Color.Blue;
            dataGridView1.Columns["TON_TP"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Regular);

            dataGridView1.Columns["BLOCK_QTY"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["BLOCK_QTY"].DefaultCellStyle.BackColor = Color.Red;
            dataGridView1.Columns["BLOCK_QTY"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Regular);

           
        }

        public void formatDataGridViewtraTON_GOP2(DataGridView dataGridView1)
        {
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;

            dataGridView1.Columns["CHO_KIEM"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["CHO_CS_CHECK"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["CHO_KIEM_RMA"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["TONG_TON_KIEM"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["BTP"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["TON_TP"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["BLOCK_QTY"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["GRAND_TOTAL_STOCK"].DefaultCellStyle.Format = "#,0";



            dataGridView1.Columns["G_NAME_KD"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["G_NAME_KD"].DefaultCellStyle.BackColor = Color.Black;
            dataGridView1.Columns["G_NAME_KD"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["GRAND_TOTAL_STOCK"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["GRAND_TOTAL_STOCK"].DefaultCellStyle.BackColor = Color.Black;
            dataGridView1.Columns["GRAND_TOTAL_STOCK"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);


            dataGridView1.Columns["TONG_TON_KIEM"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["TONG_TON_KIEM"].DefaultCellStyle.BackColor = Color.Blue;
            dataGridView1.Columns["TONG_TON_KIEM"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Regular);

            dataGridView1.Columns["BTP"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["BTP"].DefaultCellStyle.BackColor = Color.Blue;
            dataGridView1.Columns["BTP"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Regular);

            dataGridView1.Columns["TON_TP"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["TON_TP"].DefaultCellStyle.BackColor = Color.Blue;
            dataGridView1.Columns["TON_TP"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Regular);

            dataGridView1.Columns["BLOCK_QTY"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["BLOCK_QTY"].DefaultCellStyle.BackColor = Color.Red;
            dataGridView1.Columns["BLOCK_QTY"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Regular);


        }



        public void formatDataGridViewtraTON_TACH(DataGridView dataGridView1)
        {
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;

            dataGridView1.Columns["NHAPKHO"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["XUATKHO"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["TONKHO"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["BLOCK_QTY"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["GRAND_TOTAL_TP"].DefaultCellStyle.Format = "#,0";



            dataGridView1.Columns["G_CODE"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["G_CODE"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["G_CODE"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);


            dataGridView1.Columns["G_NAME"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.BackColor = Color.Black;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["NHAPKHO"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["NHAPKHO"].DefaultCellStyle.BackColor = Color.Blue;
            dataGridView1.Columns["NHAPKHO"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Regular);


            dataGridView1.Columns["XUATKHO"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["XUATKHO"].DefaultCellStyle.BackColor = Color.Blue;
            dataGridView1.Columns["XUATKHO"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Regular);

            dataGridView1.Columns["TONKHO"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["TONKHO"].DefaultCellStyle.BackColor = Color.Blue;
            dataGridView1.Columns["TONKHO"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Regular);

           

            dataGridView1.Columns["BLOCK_QTY"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["BLOCK_QTY"].DefaultCellStyle.BackColor = Color.Red;
            dataGridView1.Columns["BLOCK_QTY"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Regular);


            dataGridView1.Columns["GRAND_TOTAL_TP"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["GRAND_TOTAL_TP"].DefaultCellStyle.BackColor = Color.Black;
            dataGridView1.Columns["GRAND_TOTAL_TP"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);


            
        }



        public string generate_condition_tonkho()
        {
            string query = "WHERE 1=1";

            string chicotonkho = "";
            if(checkBox2.Checked == true)
            {
                chicotonkho = "AND THANHPHAM.TONKHO >0 ";
            }

            string code;
            if (textBox1.Text != "")
            {
                code = "AND M100.G_NAME LIKE '%" + textBox1.Text + "%' ";
            }
            else
            {
                code = "";
            }
           
            string cmscode = "";
            if (textBox2.Text != "")
            {
                cmscode = "AND M100.G_CODE='" + textBox2.Text + "'";
            }
            else
            {
                cmscode = "";
            }
            query +=  chicotonkho + code + cmscode;
            return query;
        }


        public string generate_condition_in_out_kho(string inout)
        {
            string query = "WHERE ";
            string fromdate, todate;
            fromdate = new Form1().STYMD2(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day);
            todate = new Form1().STYMD2(dateTimePicker2.Value.Year, dateTimePicker2.Value.Month, dateTimePicker2.Value.Day);
            string ngaythang = "IO_Date BETWEEN '" + fromdate + "' AND '" + todate + "' ";
            if (checkBox1.Checked == true)
            {
                ngaythang = " 1=1 ";
            }
            string code;
            if (textBox1.Text != "")
            {
                code = "AND M100.G_NAME LIKE '%" + textBox1.Text + "%' ";
            }
            else
            {
                code = "";
            }                      
            string cust_name_kd = "";
            if (textBox3.Text != "")
            {
                cust_name_kd = "AND Customer_ShortName LIKE '%" + textBox3.Text + "%' ";
            }
            else
            {
                cust_name_kd = "";
            } 
            string cmscode = "";
            if (textBox2.Text != "")
            {
                cmscode = "AND M100.G_CODE='" + textBox2.Text + "'";
            }
            else
            {
                cmscode = "";
            }

            string capbu = " AND Customer_ShortName <> 'CMSV'";            

            if(checkBox3.Checked == true || inout=="IN")
            {
                capbu = "";
            }

            string in_out_check = "";
            if(inout=="IN")
            {
                in_out_check = "AND IO_Type='IN'";
            }  
            else
            {
                in_out_check = "AND IO_Type='OUT'";
            }
            query += ngaythang + code + cust_name_kd +  cmscode + capbu + in_out_check;
            return query;
        }


        public void trakho()
        {
            if(input_flag == 1)
            {
                pictureBox1.Show();
                check_Input_Async();
            }
            else if(output_flag == 1)
            {
                pictureBox1.Show();
                check_Ouput_Async();
            }
            else if (ton_gop_flag == 1)
            {
                pictureBox1.Show();
                checkTonGop_Async();
            }
            else if (ton_gop2_flag == 1)
            {
                pictureBox1.Show();
                checkTonGop2_Async();
            }
            else if (ton_tach_flag == 1)
            {
                pictureBox1.Show();
                checkTonTach_Async();
            }
        }

        public void checkTonGop_Async()
        {
            try
            {
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();

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

                string condition = generate_condition_tonkho();
                dtgv1_data = pro.report_TONKHOFULL(condition); 
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.ToString());
            }
        }

        public void checkTonGop2_Async()
        {
            try
            {
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();

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

                string condition = generate_condition_tonkho();
                dtgv1_data = pro.report_TONKHOFULLKD(condition);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.ToString());
            }
        }

        public void checkTonTach_Async()
        {
            try
            {
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();
                string condition = generate_condition_tonkho();
                dtgv1_data = pro.report_TONKHOTACH(condition);          
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.ToString());
            }
        }

        public void check_Ouput_Async()
        {
            try
            {
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();
                string condition = generate_condition_in_out_kho("OUT");
                dtgv1_data = pro.report_TONKHO_OUTPUT(condition);              
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.ToString());
            }
        }

        public void check_Input_Async()
        {
            try
            {
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();
                string condition = generate_condition_in_out_kho("IN");
                dtgv1_data = pro.report_TONKHO_INPUT(condition);               
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.ToString());
            }
        }


        private void button3_Click(object sender, EventArgs e)
        {
            dataGridView1.Columns.Clear();
            input_flag = 0;
            output_flag = 0;
            ton_gop_flag = 1;
            ton_tach_flag = 0;
            ton_gop2_flag = 0;

            if (!backgroundWorker1.IsBusy)
            {
                backgroundWorker1.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Tiến trình khác đang chạy, thử lại sau");
            }

        }

        private void KHOTHANHPHAM_Load(object sender, EventArgs e)
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
            dataGridView1.Columns.Clear();

            input_flag = 1;
            output_flag = 0;
            ton_gop_flag = 0;
            ton_gop2_flag = 0;
            ton_tach_flag = 0;

            if (!backgroundWorker1.IsBusy)
            {
                backgroundWorker1.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Tiến trình khác đang chạy, thử lại sau");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            dataGridView1.Columns.Clear();
            input_flag = 0;
            output_flag = 0;
            ton_gop_flag = 0;
            ton_tach_flag = 1;
            ton_gop2_flag = 0;

            if (!backgroundWorker1.IsBusy)
            {
                backgroundWorker1.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Tiến trình khác đang chạy, thử lại sau");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.Columns.Clear();
            input_flag = 0;
            output_flag = 1;
            ton_gop_flag = 0;
            ton_tach_flag = 0;
            ton_gop2_flag = 0;

            if (!backgroundWorker1.IsBusy)
            {
                backgroundWorker1.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Tiến trình khác đang chạy, thử lại sau");
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            trakho();
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {

        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            pictureBox1.Hide();
            if (input_flag == 1)
            {
                dataGridView1.DataSource = dtgv1_data;
                new Form1().setRowNumber(dataGridView1);
                formatDataGridViewtraSTOCK_IN(dataGridView1);
                MessageBox.Show("Đã load: " + dtgv1_data.Rows.Count + " dòng");  
            }
            else if (output_flag == 1)
            {
                dataGridView1.DataSource = dtgv1_data;
                formatDataGridViewtraSTOCK_OUT(dataGridView1);
                new Form1().setRowNumber(dataGridView1);
                MessageBox.Show("Đã load: " + dtgv1_data.Rows.Count + " dòng");
            }
            else if (ton_gop_flag == 1)
            {

                dataGridView1.DataSource = dtgv1_data;
                formatDataGridViewtraTON_GOP(dataGridView1);
                new Form1().setRowNumber(dataGridView1);
                MessageBox.Show("Đã load: " + dtgv1_data.Rows.Count + " dòng");
            }
            else if (ton_gop2_flag == 1)
            {
                dataGridView1.DataSource = dtgv1_data;
                formatDataGridViewtraTON_GOP2(dataGridView1);
                new Form1().setRowNumber(dataGridView1);
                MessageBox.Show("Đã load: " + dtgv1_data.Rows.Count + " dòng");
            }
            else if (ton_tach_flag == 1)
            {
                dataGridView1.DataSource = dtgv1_data;
                formatDataGridViewtraTON_TACH(dataGridView1);
                new Form1().setRowNumber(dataGridView1);
                MessageBox.Show("Đã load: " + dtgv1_data.Rows.Count + " dòng");
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            dataGridView1.Columns.Clear();
            input_flag = 0;
            output_flag = 0;
            ton_gop_flag = 0;
            ton_tach_flag = 0;
            ton_gop2_flag = 1;

            if (!backgroundWorker1.IsBusy)
            {
                backgroundWorker1.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Tiến trình khác đang chạy, thử lại sau");
            }
        }
    }
}
