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
    public partial class INSPECT_INPUT : Form
    {
        public INSPECT_INPUT()
        {
            InitializeComponent();
            comboBox1.Items.Add("NM1");
            comboBox1.Items.Add("NM2");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            string EMPL_NO = textBox1.Text.ToUpper();
            string PROCESS_LOT_NO = textBox2.Text.ToUpper();
            string INPUT_DATETIME = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            string INSPECT_INPUT_QTY_EA = textBox3.Text.ToUpper();
            string INSPECT_INPUT_QTY_KG = textBox4.Text.ToUpper();
            string FACTORY = comboBox1.Text;

            string values = "('002','" + EMPL_NO + "','" + PROCESS_LOT_NO + "','" + INPUT_DATETIME + "','" + INSPECT_INPUT_QTY_EA + "','" + INSPECT_INPUT_QTY_KG + "','" + FACTORY + "')";
            if(EMPL_NO=="" || PROCESS_LOT_NO=="" || INSPECT_INPUT_QTY_EA =="" || INSPECT_INPUT_QTY_KG=="" || FACTORY=="")
            {
                MessageBox.Show("Không để trống 1 ô nào !");
            }
            else
            {
                try
                {
                    ProductBLL pro = new ProductBLL();
                    DataTable dt = new DataTable();
                    dt = pro.report_inspection_insert_input(values);
                    MessageBox.Show("NHẬP THÀNH CÔNG !");
                    dt = pro.report_inspection_all_input_data("");
                    dataGridView1.DataSource = dt;
                }
                catch(Exception ex)
                {
                    MessageBox.Show("Loi : " + ex.ToString());
                }
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            //MessageBox.Show("Text changed !");
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dt = pro.report_inspection_check_lot_no(textBox2.Text);
            if (dt.Rows.Count > 0)
            {
                foreach(DataRow row in dt.Rows)
                {
                    label7.Text = row[0].ToString();
                }

            }
            else
            {
                label7.Text = "LOT không tồn tại, hoặc lot này chưa được nhập kiểm";
            }
        }

        private void INSPECT_INPUT_Load(object sender, EventArgs e)
        {

        }
    }
}
