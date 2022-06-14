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
    public partial class QuanLyKhachHang : Form
    {
        public QuanLyKhachHang()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ProductBLL pro = new ProductBLL();
            string searchValue = "";
            searchValue = $" CUST_NAME_KD LIKE '%{textBox1.Text}%'";
            DataTable dt = pro.report_getCustomerList1(searchValue);
            dataGridView1.DataSource = dt;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ProductBLL pro = new ProductBLL();
            string CUST_CD = textBox3.Text;
            string CUST_NAME = textBox1.Text;
            string CUST_NAME_KD = textBox2.Text;
            if(CUST_CD !="" && CUST_NAME != "" && CUST_NAME_KD != "")
            {
                try
                {
                    string addValue = $"('002','{CUST_CD}','{CUST_NAME}','{CUST_NAME_KD}')";
                    DataTable dt = pro.report_addCustomer(addValue);
                    MessageBox.Show("Thêm khách thành công");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi: " + ex.ToString());
                }
               
            }
            else
            {
                MessageBox.Show("Nhập thông tin khách đầy đủ vào, có mỗi 3 ô cũng ko nhập được hết !");
            }
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ProductBLL pro = new ProductBLL();
            string CUST_CD = textBox3.Text;
            string CUST_NAME = textBox1.Text;
            string CUST_NAME_KD = textBox2.Text;

            if (CUST_CD != "" && CUST_NAME != "" && CUST_NAME_KD != "")
            {
                try
                {
                    string updateValue = $"SET CUST_NAME='{CUST_NAME}', CUST_NAME_KD='{CUST_NAME_KD}' WHERE CUST_CD='{CUST_CD}'";
                    DataTable dt = pro.report_updateCustomer(updateValue);
                    MessageBox.Show("Update thông tin khách thành công");
                }
                catch(Exception ex)
                {
                    MessageBox.Show("Lỗi: " + ex.ToString());
                }
               
            }
            else
            {
                MessageBox.Show("Nhập thông tin khách đầy đủ vào, có mỗi 3 ô cũng ko nhập được hết !");
            }
        }

        private void dataGridView1_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewRow row = dataGridView1.Rows[e.RowIndex];
            textBox2.Text = row.Cells["CUST_NAME"].Value.ToString();
            textBox1.Text = row.Cells["CUST_NAME_KD"].Value.ToString();
            textBox3.Text = row.Cells["CUST_CD"].Value.ToString();
        }
    }
}
