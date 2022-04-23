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
    public partial class InvoiceForm : Form
    {
        public InvoiceForm()
        {
            InitializeComponent();
        }

        public string loginIDInvoiceForm = "";
        public string CUST_NAME = "", G_NAME = "", CUST_CD = "", EMPL_NO = "", G_CODE = "", PO_NO = "", DELIVERY_QTY = "", DELIVERY_DATE = "", NOCANCEL = "", REMARK = "", DELIVERY_ID = "";

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //update invoice button
        private void button4_Click(object sender, EventArgs e)
        {
               
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            Form1 frm1 = new Form1();
            if (textBox3.Text == "" || textBox4.Text == "" || textBox6.Text == "" || comboBox1.Text == "" || comboBox2.Text == "" || comboBox3.Text == "" || comboBox4.Text == "")
            {
                MessageBox.Show("Không để trống 1 ô nào !");
            }
            else
            {
                try
                {
                    string CUST_CD, EMPL_NO1, G_CODE, PO_NO, DELIVERY_QTY1, DELIVERY_DATE, NOCANCEL, REMARK, ID_DELIVERY;
                    string CTR_CD = "002";
                    G_CODE = comboBox3.Text;
                    DELIVERY_QTY1 = textBox3.Text;
                    CUST_CD = comboBox4.Text;
                    EMPL_NO1 = EMPL_NO;
                    DELIVERY_DATE = textBox4.Text;
                    PO_NO = textBox6.Text;
                    REMARK = textBox8.Text;
                    NOCANCEL = "1";
                    ID_DELIVERY = DELIVERY_ID;
                    int po_balance = pro.checkPOBalance(CUST_CD, G_CODE, PO_NO);                    
                    int po_balance_beforechange = po_balance + int.Parse(DELIVERY_QTY);
                    MessageBox.Show("PO Balance before change = " + po_balance_beforechange);
                    DateTime dlidate = DateTime.Parse(DELIVERY_DATE);
                    int check_date = new Form1().checkDate(dlidate);
                    DateTime podate = DateTime.Parse(pro.getPODDate(G_CODE, PO_NO)); // get podate and convert to datetime
                    int checkinvoicedatevspodate = new Form1().checkInvoicevsPODate(dlidate, podate);
            
                    
                    if(check_date == 0)
                    {
                        MessageBox.Show("Ngày invoice không được lớn hơn ngày hôm nay ");
                    }
                    else if (checkinvoicedatevspodate == 0)
                    {
                        MessageBox.Show("Ngày invoice không được nhỏ hơn ngày PO ");
                    }
                    else if(int.Parse(DELIVERY_QTY1) > po_balance_beforechange && check_date == 1)
                    {
                        MessageBox.Show("Giao hàng nhiều hơn số lượng PO ! Update Invoice mới thất bại !");
                    }
                    else if (int.Parse(DELIVERY_QTY1) <= po_balance_beforechange && check_date == 1)
                    {
                        pro.UpdateInvoice(CTR_CD, CUST_CD, EMPL_NO1, G_CODE, PO_NO, DELIVERY_QTY1, DELIVERY_DATE, NOCANCEL, ID_DELIVERY);
                        pro.writeHistory("002", loginIDInvoiceForm, "DELIVERY TABLE", "SUA", "THEM INVOICE CODE: " + G_CODE + " , QTY = " + DELIVERY_QTY + ", PO NO: " + PO_NO, "" + ID_DELIVERY);
                        MessageBox.Show("Đã update Invoice " + DELIVERY_ID + " thành công !");
                        this.Close();
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }

        public void updateform()
        {
            comboBox1.Text = G_NAME;
            comboBox3.Text = G_CODE;
            comboBox2.Text = CUST_NAME;
            comboBox4.Text = CUST_CD;
            textBox3.Text = DELIVERY_QTY;
            textBox4.Text = DELIVERY_DATE;
            textBox6.Text = PO_NO;
            textBox8.Text = REMARK;

        }
        private void InvoiceForm_Load(object sender, EventArgs e)
        {
            

        }

        private void textBox9_TextChanged(object sender, EventArgs e)
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

        private void button1_Click(object sender, EventArgs e)
        {
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            Form1 frm1 = new Form1();
            if ( textBox3.Text == "" || textBox4.Text == ""  || textBox6.Text == ""  || comboBox1.Text == "" || comboBox2.Text == "" || comboBox3.Text == "" || comboBox4.Text == "")
            {
                MessageBox.Show("Không để trống 1 ô nào !");
            }
            else
            {
                try
                {
                    string CUST_CD, EMPL_NO, G_CODE, PO_NO, DELIVERY_QTY, DELIVERY_DATE, NOCANCEL, REMARK;
                    string CTR_CD = "002";
                    G_CODE = comboBox3.Text;
                    DELIVERY_QTY = textBox3.Text;
                    CUST_CD = comboBox4.Text;
                    EMPL_NO = loginIDInvoiceForm;
                    DELIVERY_DATE = textBox4.Text;                    
                    PO_NO = textBox6.Text;                    
                    REMARK = textBox8.Text;
                    NOCANCEL = "1";
                    int po_balance = pro.checkPOBalance(CUST_CD, G_CODE, PO_NO);
                    DateTime dlidate = DateTime.Parse(DELIVERY_DATE);
                    int check_date = new Form1().checkDate(dlidate);
                    DateTime podate = DateTime.Parse(pro.getPODDate(G_CODE, PO_NO)); // get podate and convert to datetime
                    int checkinvoicedatevspodate = new Form1().checkInvoicevsPODate(dlidate, podate);

                    
                    if (check_date == 0)
                    {
                        MessageBox.Show("Ngày invoice không được lớn hơn ngày hôm nay ");
                    }
                    else if(checkinvoicedatevspodate==0)
                    {
                        MessageBox.Show("Ngày invoice không được nhỏ hơn ngày PO ");
                    }
                    else if(int.Parse(DELIVERY_QTY) > po_balance)
                    {
                        MessageBox.Show("Giao hàng nhiều hơn số lượng PO ! Thêm Invoice mới thất bại !");                        
                    }
                    else if (int.Parse(DELIVERY_QTY) <= po_balance)
                    {
                        pro.InsertInvoice(CTR_CD, CUST_CD, EMPL_NO, G_CODE, PO_NO, DELIVERY_QTY, DELIVERY_DATE, NOCANCEL);
                        pro.writeHistory("002", EMPL_NO, "DELIVERY TABLE", "THEM", "THEM INVOICE CODE: " + G_CODE + " , QTY = " + DELIVERY_QTY + ", PO NO: " + PO_NO, "0");

                        MessageBox.Show("Đã thêm Invoice mới thành công !");
                        this.Close();

                    }


                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }
    }
}
