using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoClick
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            MessageBox.Show(loginIDfrm3);       
            
        }

        private void comboBox1_TextUpdate(object sender, EventArgs e)
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

        private void Form3_Load(object sender, EventArgs e)
        {
            comboBox5.Items.Add("Thông Thường");
            comboBox5.Items.Add("SDI");
            comboBox5.Items.Add("ETC");
            comboBox5.Items.Add("SAMPLE");

            comboBox6.Items.Add("GC");
            comboBox6.Items.Add("SK");
            comboBox6.Items.Add("KD");
            comboBox6.Items.Add("VN");
            comboBox6.Items.Add("Sample");
            comboBox6.Items.Add("Vai bac 4");
            comboBox6.Items.Add("ETC");
           // button3.Enabled = false;
            

        }
        public string loginIDfrm3 = "";

        private void button1_Click(object sender, EventArgs e)
        {
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            Form1 frm1 = new Form1();
            if (textBox1.Text == null || textBox1.Text == "" || textBox2.Text == "" || textBox3.Text == "" || textBox4.Text == "" || textBox5.Text == "" || comboBox1.Text =="" || comboBox2.Text == "" || comboBox3.Text == "" || comboBox4.Text == "" || comboBox5.Text == "" || comboBox6.Text == "")
            {
                MessageBox.Show("Không để trống 1 ô nào !");
            }
            else
            {
                try
                {           
                    string PROD_REQUEST_DATE, CODE_50="", CODE_55="", G_CODE, RIV_NO, PROD_REQUEST_QTY, CUST_CD, EMPL_NO, REMK, DELIVERY_DT;
                    string CTR_CD = "002";
                    string CODE_03 = "01";
                    String ngaygiohethong = pro.getsystemDateTime();
                    String ngayhethong = ngaygiohethong.Substring(0, 4) + ngaygiohethong.Substring(5, 2) + ngaygiohethong.Substring(8, 2);
                    PROD_REQUEST_DATE = ngayhethong;
                    switch (comboBox5.Text)
                    {
                        case "Thông Thường":
                            CODE_55 = "01";
                            break;
                        case "SDI":
                            CODE_55 = "02";
                            break;
                        case "ETC":
                            CODE_55 = "03";
                            break;
                        case "SAMPLE":
                            CODE_55 = "04";
                            break;
                        default:
                            break;                           

                    }
                    switch (comboBox6.Text)
                    {
                        case "GC":
                            CODE_50 = "01";
                            break;
                        case "SK":
                            CODE_50 = "02";
                            break;
                        case "KD":
                            CODE_50 = "03";
                            break;
                        case "VN":
                            CODE_50 = "04";
                            break;
                        case "SAMPLE":
                            CODE_50 = "05";
                            break;
                        case "Vai bac 4":
                            CODE_50 = "06";
                            break;
                        case "ETC":
                            CODE_50 = "07";
                            break;
                        default:
                            break;

                    }

                    G_CODE = comboBox3.Text;
                    RIV_NO = G_CODE.Substring(7);
                    PROD_REQUEST_QTY = textBox3.Text;
                    CUST_CD = comboBox4.Text;
                    EMPL_NO = loginIDfrm3;
                    REMK = textBox4.Text;
                    DELIVERY_DT = textBox5.Text;                 
                               
                    dt = pro.getLastYCSXNo();
                    string lastycsxno = "";
                    if (dt.Rows.Count > 0)
                    {
                        int yccuoiint = 0;
                        lastycsxno = dt.Rows[0]["PROD_REQUEST_NO"].ToString();
                        if (lastycsxno.Substring(0, 3) != frm1.CreateHeader2())
                        {
                            yccuoiint = 0;
                        }
                        else
                        {
                            lastycsxno = dt.Rows[0]["PROD_REQUEST_NO"].ToString().Substring(3, 4);
                            yccuoiint = int.Parse(lastycsxno);
                        }

                        yccuoiint = yccuoiint + 1;
                        if (yccuoiint < 10)
                        {
                            lastycsxno = "000" + yccuoiint;
                        }
                        else if (yccuoiint < 100)
                        {
                            lastycsxno = "00" + yccuoiint;
                        }
                        else if (yccuoiint < 1000)
                        {
                            lastycsxno = "0" + yccuoiint;
                        }
                        else
                        {
                            lastycsxno = "" + yccuoiint;
                        }
                        string PROD_REQUEST_NO = frm1.CreateHeader2() + lastycsxno;

                        /*
                        int check_riv = pro.checkRIV_NO(G_CODE, RIV_NO);
                        string checkUSEYN = pro.checkM100UseYN(G_CODE);

                        if (check_riv != 1)
                        {
                            MessageBox.Show("Code " + G_CODE + " không tồn tại REVISION trong bảng BOM, check lại REVISION hoặc liên hệ RND");
                        }
                        else if(checkUSEYN=="N")
                        {
                            MessageBox.Show("Code " + G_CODE + " đã bị khóa, có thể ver này không còn được sử dụng");
                        }
                        else
                        {
                            pro.InsertYCSX(CTR_CD, PROD_REQUEST_DATE, PROD_REQUEST_NO, CODE_50, CODE_03, CODE_55, G_CODE, RIV_NO, PROD_REQUEST_QTY, CUST_CD, EMPL_NO, REMK, EMPL_NO, EMPL_NO, DELIVERY_DT);
                            pro.writeHistory(CTR_CD, EMPL_NO, "YCSX TABLE", "THEM", "THEM YCSX", "0");
                        }

                        */
                        
                        string checkUSEYN = pro.checkM100UseYN(G_CODE);

                       
                        if (checkUSEYN == "N")
                        {
                            MessageBox.Show("Code " + G_CODE + " đã bị khóa, có thể ver này không còn được sử dụng");
                        }
                        else
                        {
                            pro.InsertYCSX(CTR_CD, PROD_REQUEST_DATE, PROD_REQUEST_NO, CODE_50, CODE_03, CODE_55, G_CODE, "A", PROD_REQUEST_QTY, CUST_CD, EMPL_NO, REMK, EMPL_NO, EMPL_NO, DELIVERY_DT);
                            pro.writeHistory(CTR_CD, EMPL_NO, "YCSX TABLE", "THEM", "THEM YCSX", "0");
                        }


                    }
                    else
                    {
                        MessageBox.Show("Loi");
                    }



                    

                    MessageBox.Show("Đã hoàn thành yêu cầu sản xuất hàng loạt");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
