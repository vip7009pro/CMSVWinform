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
    public partial class PQCForm2 : Form
    {
        public PQCForm2()
        {
            InitializeComponent();
        }

        public String format_time(String time)
        {
            String formated_time = "";
            formated_time = time.Substring(0, 2) + ":" + time.Substring(2, 2) + ":00";
            return formated_time;
        }
        public void insert_lineqc_checksheet()
        {
            ProductBLL pro = new ProductBLL();        


            //mang ket qua
            String[] checksheet_time_array = { "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""};
            String[] checksheet_result_array = { "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""};

            //lay thong tin co ban ve code
            string EMPL_NO = textBox2.Text.ToUpper();
            string PROCESS_LOT_NO = textBox1.Text.ToUpper();
            string REMARK = textBox10.Text;


            // lay ngay gio phut giay san xuat dua tren lot sx
            String prod_datetime = pro.report_inspection_check_prod_date(PROCESS_LOT_NO);
            //tach lay ngay sx
            String prod_date = prod_datetime.Substring(0, 10);
            

            //lay data checksheet từ datagridview2
            for (int r = 0; r < dataGridView2.Rows.Count - 1; r++)
            {
                DataGridViewRow row = dataGridView2.Rows[r];

                if(row.Cells[0].Value.ToString()=="" || row.Cells[1].Value.ToString() == "")
                {
                    MessageBox.Show("Không được bỏ trống thời gian hoặc kết quả nào !");
                }
                else
                {
                    checksheet_time_array[r] = prod_date + " " + format_time(row.Cells[0].Value.ToString());

                    if (int.Parse(row.Cells[1].Value.ToString()) == 0)
                    {
                        checksheet_result_array[r] = "NG";
                    }
                    else if (int.Parse(row.Cells[1].Value.ToString()) == 1)
                    {
                        checksheet_result_array[r] = "OK";
                    }
                    else
                    {
                        MessageBox.Show("Kết quả chỉ nhập  1 or 0");
                    }
                }               
                
            }

            /*
             * test show data array
            for (int r = 0; r < dataGridView2.Rows.Count - 1; r++)
            {
                MessageBox.Show(checksheet_time_array[r] + " :    " + checksheet_result_array[r]);
            }
            */


            //ghep chuoi time va ket qua
            String time_arr = "'";
            String result_arr = "'";
            for(int k = 0;k<15;k++)
            {
                time_arr += checksheet_time_array[k] + "','";
                result_arr += checksheet_result_array[k] + "','";
            }
            time_arr = time_arr.Substring(0, time_arr.Length - 2);
            result_arr = result_arr.Substring(0, result_arr.Length - 2);

            //MessageBox.Show(time_arr + "\n KQ: " + result_arr);

            //ghep chuoi ket qua
            String values = "('002','" + PROCESS_LOT_NO + "','" + EMPL_NO + "'," + time_arr + "," + result_arr + ",'" + REMARK + "')";
            MessageBox.Show(values);
            //insert giá trị vào bảng
            pro.pqc2_insert(values);
            MessageBox.Show("Nhập thành công");


        }
        private void button1_Click(object sender, EventArgs e)
        {
            if(textBox1.Text =="" || textBox2.Text == "" || dataGridView2.Rows.Count <=1)
            {
                MessageBox.Show("Nhập data cho đầy đủ vào !");
            }
            else
            {     
                try
                {
                    insert_lineqc_checksheet();
                }
                catch(Exception ex)
                {
                    MessageBox.Show("Loi : " + ex.ToString());
                }
                
            }                        
        }

        private void PQCForm2_Load(object sender, EventArgs e)
        {
            dataGridView2.Columns.Add("TIME", "Thời gian check");
            dataGridView2.Columns.Add("RESULT", "Kết quả check");
            textBox1.Text = "DC1DS087";
            textBox2.Text = "NHU1903";
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            String lot_info = "CODE: ";
            dt = pro.report_inspection_check_lot_no(textBox1.Text);
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    lot_info += row[0].ToString() + "\n NGÀY SẢN XUẤT: " +  pro.report_inspection_check_prod_date(textBox1.Text); 
                }
                label7.Text = lot_info;

            }            
            else
            {
                label7.Text = "LOT không tồn tại, hoặc lot này chưa được nhập kiểm";
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            ProductBLL pro = new ProductBLL();
            label3.Text = "LINEQC: " +  pro.report_inspection_check_EMPLNAME(textBox2.Text);
                 
        }
    }
}
