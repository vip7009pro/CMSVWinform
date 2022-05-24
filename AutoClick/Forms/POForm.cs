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
    public partial class POForm : Form
    {
        public POForm()
        {
            InitializeComponent();
        }

        public string loginIDpoForm = "";
        public string CUST_NAME = "", G_NAME = "", CUST_CD = "", EMPL_NO = "", G_CODE = "", PO_NO = "", DELIVERY_QTY = "", DELIVERY_DATE = "", NOCANCEL = "", REMARK = "", PO_QTY = "", RD_DATE = "", PROD_PRICE= "", PO_DATE = "", PO_ID="";

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void POForm_Load(object sender, EventArgs e)
        {

        }


        //update PO button

        public void updatePO()
        {
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            Form1 frm1 = new Form1();
            if (textBox3.Text == "" || textBox4.Text == "" || textBox5.Text == "" || textBox6.Text == "" || textBox7.Text == "" || comboBox1.Text == "" || comboBox2.Text == "" || comboBox3.Text == "" || comboBox4.Text == "")
            {
                MessageBox.Show("Không để trống 1 ô nào !");
            }
            else
            {
                try
                {
                    string CUST_CD, EMPL_NO1, G_CODE, PO_NO, PO_QTY, PO_DATE, RD_DATE, PROD_PRICE, REMARK, ID_PO;
                    string CTR_CD = "002";
                    G_CODE = comboBox3.Text;
                    PO_QTY = textBox3.Text;
                    CUST_CD = comboBox4.Text;
                    EMPL_NO1 = EMPL_NO;
                    PO_DATE = textBox4.Text;
                    RD_DATE = textBox5.Text;
                    PO_NO = textBox6.Text;
                    PROD_PRICE = textBox7.Text;
                    REMARK = textBox8.Text;
                    ID_PO = PO_ID;
                    DateTime podate = DateTime.Parse(PO_DATE);
                    int check_date = new Form1().checkDate(podate);
                    int delivered_qty = pro.checkDeliveredQTy(CUST_CD, G_CODE, PO_NO);
                    int po_qty = int.Parse(PO_QTY);

                    if (check_date == 0)
                    {
                        MessageBox.Show("Ngày PO không được lớn hơn ngày hôm nay ");
                    }
                    else if(po_qty < delivered_qty)
                    {
                        MessageBox.Show("Số lượng PO đã sửa nhỏ hơn số lượng đã giao hàng, k đc nhé ");
                    }
                    else if (1==1 /*pro.checkPOExist(CUST_CD, G_CODE, PO_NO) != -1*/)
                    {
                        pro.UpdatePO(CTR_CD, CUST_CD, EMPL_NO1, G_CODE, PO_NO, PO_QTY, PO_DATE, RD_DATE, PROD_PRICE, ID_PO);
                        pro.writeHistory(CTR_CD, EMPL_NO, "PO TABLE", "SUA", "SUA PO CODE: " + G_CODE + ", QTY = " + PO_QTY + " , MA KHACH: " + CUST_CD, PO_ID);
                        MessageBox.Show("UPDATE xong PO " + PO_ID + " r nhé thím ! !");
                        this.Close();
                    }
                    else
                    {
                        MessageBox.Show("Check lại thông tin, làm méo gì có PO nào thông tin như vậy mà update ?! ĐKM !");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
            updatePO();
        }

        public void updateform()
        {
            comboBox1.Text = G_NAME;
            comboBox3.Text = G_CODE;
            comboBox2.Text = CUST_NAME;
            comboBox4.Text = CUST_CD;
            textBox3.Text = PO_QTY;
            textBox4.Text = PO_DATE;
            textBox5.Text = RD_DATE;
            textBox6.Text = PO_NO;
            textBox7.Text = PROD_PRICE;
            textBox8.Text = REMARK;
        }


        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dt = pro.getnameandcode(textBox1.Text);
            comboBox1.DataSource = dt;
            comboBox1.DisplayMember = "G_NAME";
            comboBox1.ValueMember = "G_NAME";
            comboBox3.DataSource = dt;
            comboBox3.DisplayMember = "G_CODE";
            comboBox3.ValueMember = "G_CODE";
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dt = pro.getcustomerinfo(textBox2.Text);
            comboBox2.DataSource = dt;
            comboBox2.DisplayMember = "CUST_NAME";
            comboBox2.ValueMember = "CUST_NAME";
            comboBox4.DataSource = dt;
            comboBox4.DisplayMember = "CUST_CD";
            comboBox4.ValueMember = "CUST_CD";
        }

        public void insertPO()
        {
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            Form1 frm1 = new Form1();
            if (textBox1.Text == null || textBox1.Text == "" || textBox2.Text == "" || textBox3.Text == "" || textBox4.Text == "" || textBox5.Text == "" || textBox6.Text == "" || textBox7.Text == "" || textBox8.Text == "" || comboBox1.Text == "" || comboBox2.Text == "" || comboBox3.Text == "" || comboBox4.Text == "")
            {
                MessageBox.Show("Không để trống 1 ô nào !");
            }
            else
            {
                try
                {
                    string CUST_CD, EMPL_NO, G_CODE, PO_NO, PO_QTY, PO_DATE, RD_DATE, PROD_PRICE, REMARK;
                    string CTR_CD = "002";
                    G_CODE = comboBox3.Text;
                    PO_QTY = textBox3.Text;
                    CUST_CD = comboBox4.Text;
                    EMPL_NO = loginIDpoForm;
                    PO_DATE = textBox4.Text;
                    RD_DATE = textBox5.Text;
                    PO_NO = textBox6.Text;
                    PROD_PRICE = textBox7.Text;
                    REMARK = textBox8.Text;
                    DateTime podate = DateTime.Parse(PO_DATE);
                    int check_date = new Form1().checkDate(podate);

                    if (check_date == 0)
                    {
                        MessageBox.Show("Ngày PO không được lớn hơn ngày hôm nay ");
                    }
                    else if (pro.checkPOExist(CUST_CD, G_CODE, PO_NO) != -1)
                    {
                        MessageBox.Show("Đã tồn tại PO, thêm PO thất bại!");
                    }
                    else
                    {
                        pro.InsertPO(CTR_CD, CUST_CD, EMPL_NO, G_CODE, PO_NO, PO_QTY, PO_DATE, RD_DATE, PROD_PRICE);
                        pro.writeHistory(CTR_CD, EMPL_NO, "PO TABLE", "THEM", "THEM PO CODE: " + G_CODE + ", QTY = " + PO_QTY + " , MA KHACH: " + CUST_CD, "0");
                        MessageBox.Show("Đã thêm PO mới thành công !");
                        this.Close();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            insertPO();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
