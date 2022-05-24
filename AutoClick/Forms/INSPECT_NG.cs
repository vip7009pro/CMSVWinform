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
    public partial class INSPECT_NG : Form
    {
        public INSPECT_NG()
        {
            InitializeComponent();
        }

      

        private void INSPECT_NG_Load(object sender, EventArgs e)
        {
            dataGridView2.Columns.Add("ERR_CODE", "Mã Lỗi");
            dataGridView2.Columns.Add("ERR_QTY", "Số lượng Lỗi");
            comboBox1.Items.Add("NM1");
            comboBox1.Items.Add("NM2");
            comboBox1.Text = "NM1";
            checkBox1.Checked = true;
        }

        public string STYMD2(int y, int m, int d)
        {
            string ymd, sty, stm, std;
            sty = "" + y;
            stm = "" + m;
            std = "" + d;
            if (m < 10)
            {
                stm = "0" + m;
            }
            if (d < 10)
            {
                std = "0" + d;
            }
            ymd = sty + "-" + stm + "-" + std;
            return ymd;
        }
        public string STYMD2_plus(int y, int m, int d)
        {
            int yy, mm, dd;
            yy = y; mm = m; dd = d;
            DateTime curr_day = new DateTime(yy, mm, dd);
            DateTime next_day = curr_day.AddDays(1);

            string ymd, sty, stm, std;
            sty = "" + next_day.Year;
            stm = "" + next_day.Month;
            std = "" + next_day.Day;

            if (m < 10)
            {
                stm = "0" + m;
            }
            if (d < 10)
            {
                std = "0" + d;
            }
            ymd = sty + "-" + stm + "-" + std;
            return ymd;
        }


        public void insertNGdata()
        {
            
            try
            {
                int[] err_array = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
                string EMPL_NO = textBox1.Text.ToUpper();
                string PROCESS_LOT_NO = textBox2.Text.ToUpper();
                string INSPECT_DATETIME = textBox3.Text.ToUpper();
                string FACTORY = comboBox1.Text.ToUpper();
                string LINEQC_PIC = textBox5.Text.ToUpper();
                string MACHINE_NO = textBox6.Text.ToUpper();
                string INSPECT_TOTAL_QTY = textBox7.Text.ToUpper();
                string INSPECT_OK_QTY = textBox8.Text.ToUpper();
                string INSPECT_START_DATETIME = "";
                string INSPECT_FINISH_DATETIME = "";


                string inspect_start_time = textBox9.Text.ToUpper();
                string inspect_finish_time = textBox10.Text.ToUpper();
                string inspect_start_time_h = inspect_start_time.Substring(0,2);
                string inspect_start_time_m = inspect_start_time.Substring(2,2);

                string inspect_finish_time_h = inspect_finish_time.Substring(0, 2);
                string inspect_finish_time_m = inspect_finish_time.Substring(2, 2);
                inspect_start_time = inspect_start_time_h + ":" + inspect_start_time_h;
                inspect_finish_time = inspect_finish_time_h + ":" + inspect_finish_time_m;


                int inspect_start_hour = int.Parse(inspect_start_time.Substring(0, 2));
                int inspect_finish_hour = int.Parse(inspect_finish_time.Substring(0, 2));

                int inspect_year, inspect_month, inspect_day;
                inspect_year = int.Parse(INSPECT_DATETIME.Substring(0, 4));
                inspect_month = int.Parse(INSPECT_DATETIME.Substring(5, 2));
                inspect_day = int.Parse(INSPECT_DATETIME.Substring(8, 2));


                if (inspect_finish_hour >= inspect_start_hour)
                {
                    INSPECT_START_DATETIME = STYMD2(inspect_year, inspect_month, inspect_day) + " " + inspect_start_time;
                    INSPECT_FINISH_DATETIME = STYMD2(inspect_year, inspect_month, inspect_day) + " " + inspect_finish_time;
                }
                else
                {
                    INSPECT_START_DATETIME = STYMD2(inspect_year, inspect_month, inspect_day) + " " + inspect_start_time;
                    INSPECT_FINISH_DATETIME = STYMD2_plus(inspect_year, inspect_month, inspect_day) + " " + inspect_finish_time;
                }
                             

                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();

                //dataGridView2.Rows.Clear();

                string values = "";
                string err_qty = "";

                for (int r = 0; r < dataGridView2.Rows.Count - 1; r++)
                {
                    DataGridViewRow row = dataGridView2.Rows[r];  
                    int ERC = int.Parse(row.Cells[0].Value.ToString())-1;
                    if (ERC < 32 && ERC >= 0)
                    {
                        err_array[ERC] = int.Parse(row.Cells[1].Value.ToString());
                    }
                    else
                    {
                        MessageBox.Show("Mã lỗi chỉ từ 1 đến 32, nhập lại mã lỗi");
                    }

                }
                for (int i = 0; i < 31; i++)
                {
                    err_qty = err_qty + err_array[i] + "','";
                }
                err_qty = err_qty + err_array[31] + "')";

                values = "('002','" + INSPECT_START_DATETIME + "','" + INSPECT_FINISH_DATETIME + "','" + EMPL_NO + "','" + PROCESS_LOT_NO + "','" + INSPECT_DATETIME + "','" + FACTORY + "','" + LINEQC_PIC + "','" + MACHINE_NO + "','" + INSPECT_TOTAL_QTY + "','" + INSPECT_OK_QTY + "','";

                values = values + err_qty;

                //MessageBox.Show(values);

                int total_inspect_input = int.Parse(INSPECT_TOTAL_QTY);
                int total_ok_qty = int.Parse(INSPECT_OK_QTY);
                int err2toerr32_total_qty = 0;
                for(int iii=1;iii<31;iii++)
                {
                    err2toerr32_total_qty += err_array[iii];
                }
                int total_real_qty = total_ok_qty + err2toerr32_total_qty;
                if(total_inspect_input != total_real_qty)
                {
                    MessageBox.Show("Số lượng nhập vào không khớp, quy tắc : TỔNG VÀO = OK + ERR1 -> ERR31");
                }
                else
                {
                    dt = pro.report_inspection_insert_NG(values);
                    MessageBox.Show("Input thành công !");
                    dataGridView2.Rows.Clear();
                    textBox1.Text = "";
                    textBox2.Text = "";
                    textBox3.Text = "";
                    textBox4.Text = "";
                    textBox5.Text = "";
                    textBox6.Text = "";
                    textBox7.Text = "";
                    textBox8.Text = "";
                    textBox9.Text = "";
                    textBox10.Text = "";
                }
              
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi : " + ex.ToString());
            }
        }


        public void traNGData()
        {
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dt = pro.report_inspection_tra_NG();
            dataGridView1.DataSource = dt;
            
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if(textBox1.Text == "" || textBox2.Text =="" || textBox3.Text == "" || textBox5.Text == "" || textBox6.Text == "" || textBox7.Text == "" || textBox8.Text == "" || textBox9.Text == "" || textBox10.Text == "" || textBox9.Text.Length!= 4 || textBox10.Text.Length != 4)
            {
                MessageBox.Show("Data nhập thiếu hoặc chưa đúng định dạng !");
            }
            else
            {
                insertNGdata();
                traNGData();
            }

            
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dt = pro.report_inspection_check_lot_no(textBox2.Text);
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    label7.Text = row[0].ToString();
                }

            }
            else
            {
                label7.Text = "LOT không tồn tại, hoặc lot này chưa được nhập kiểm";
            }
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            if(checkBox1.Checked && textBox4.Text.Length==34)
            {
                textBox1.Text = textBox4.Text.Substring(20, 7);
                textBox2.Text = textBox4.Text.Substring(2, 8);

                string year = DateTime.Now.Year.ToString();
                string month = textBox4.Text.Substring(10, 2);
                string day = textBox4.Text.Substring(12, 2);
                string hour = textBox4.Text.Substring(14, 2);
                string minute = textBox4.Text.Substring(16, 2);
                string second = textBox4.Text.Substring(18, 2);
                textBox3.Text = year + "-" + month + "-" + day + " " + hour + ":" + minute + ":" + second;
                

                //textBox3.Text = textBox4.Text.Substring(10, 9);



                textBox5.Text = textBox4.Text.Substring(27, 7);
                textBox6.Text = textBox4.Text.Substring(0, 2);
            }
        }

        private void gradientPanel2_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
