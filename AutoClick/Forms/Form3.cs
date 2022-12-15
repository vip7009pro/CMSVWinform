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

        public void addnewYCSX()
        {
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            Form1 frm1 = new Form1();

            if (textBox1.Text == null || textBox1.Text == "" || textBox2.Text == "" || textBox3.Text == "" || textBox4.Text == "" || textBox5.Text == "" || comboBox1.Text == "" || comboBox2.Text == "" || comboBox3.Text == "" || comboBox4.Text == "" || comboBox5.Text == "" || comboBox6.Text == "")
            {
                MessageBox.Show("Không để trống 1 ô nào !");
            }
            else
            {
                try
                {
                    string PROD_REQUEST_DATE, CODE_50 = "", CODE_55 = "", G_CODE, RIV_NO, PROD_REQUEST_QTY, CUST_CD, EMPL_NO, REMK, DELIVERY_DT;
                    string CTR_CD = "002";
                    string CODE_03 = "01";
                    int TP=0, BTP = 0, CK = 0, BLOCK_QTY=0, W1 = 0, W2 = 0, W3 = 0, W4 = 0, W5 = 0, W6 = 0, W7 = 0, W8 = 0, PO_BALANCE=0, TOTAL_FCST=0, PDuyet =0;
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
                        string checkUSEYN = pro.checkM100UseYN(G_CODE);


                        DataTable tonkho_gcode = new DataTable();
                        DataTable fcst_gcode = new DataTable();
                        DataTable pobalance_gcode = new DataTable();
                        tonkho_gcode = pro.checktonkhofull_gcode(G_CODE);
                        fcst_gcode = pro.checkfcst_gcode(G_CODE);
                        pobalance_gcode = pro.checkpobalance_gcode(G_CODE);


                        if (checkUSEYN == "N")
                        {
                            MessageBox.Show("Code " + G_CODE + " đã bị khóa, có thể ver này không còn được sử dụng");
                        }
                        else
                        {
                            if(tonkho_gcode.Rows.Count>0)
                            {
                                TP = int.Parse(tonkho_gcode.Rows[0]["TON_TP"].ToString());
                                BTP = int.Parse(tonkho_gcode.Rows[0]["BTP"].ToString());
                                CK = int.Parse(tonkho_gcode.Rows[0]["TONG_TON_KIEM"].ToString());
                                BLOCK_QTY = int.Parse(tonkho_gcode.Rows[0]["BLOCK_QTY"].ToString());
                            }
                            if(fcst_gcode.Rows.Count>0)
                            {
                                W1 = int.Parse(fcst_gcode.Rows[0]["W1"].ToString());
                                W2 = int.Parse(fcst_gcode.Rows[0]["W2"].ToString());
                                W3 = int.Parse(fcst_gcode.Rows[0]["W3"].ToString());
                                W4 = int.Parse(fcst_gcode.Rows[0]["W4"].ToString());
                                W5 = int.Parse(fcst_gcode.Rows[0]["W5"].ToString());
                                W6 = int.Parse(fcst_gcode.Rows[0]["W6"].ToString());
                                W7 = int.Parse(fcst_gcode.Rows[0]["W7"].ToString());
                                W8 = int.Parse(fcst_gcode.Rows[0]["W8"].ToString());
                                TOTAL_FCST = W1 + W2 + W3 + W4 + W5 + W6 + W7 + W8;

                            }
                            if(CODE_55=="04") PDuyet = 1;
                            if (pobalance_gcode.Rows.Count>0)
                            {
                                
                                PO_BALANCE = int.Parse(pobalance_gcode.Rows[0]["PO_BALANCE"].ToString());
                                if (PO_BALANCE > 0) PDuyet = 1;
                            }

                            if (checkBox1.Checked)
                            {
                                string process_in_date = "", current_process_in_no = "", next_process_in_no = "";
                                String sDate = DateTime.Now.ToString();
                                DateTime datevalue = (Convert.ToDateTime(sDate.ToString()));
                                int dy = datevalue.Day;
                                int mn = datevalue.Month;
                                int yy = datevalue.Year;
                                string in_date = new Form1().STYMD(yy, mn, dy);                               
                                dt = pro.checkProcessInNoP500(in_date, comboBox7.Text);

                                if (dt.Rows.Count > 0)
                                {
                                    //MessageBox.Show(dt.Rows[0]["PROCESS_IN_NO"].ToString());
                                    current_process_in_no = dt.Rows[0]["PROCESS_IN_NO"].ToString();
                                    next_process_in_no = String.Format("{0:000}", int.Parse(current_process_in_no) + 1);
                                }
                                else
                                {
                                    next_process_in_no = "001";
                                }
                                
                                string insertvalueP500 = $"('002','{in_date}','{next_process_in_no}','{next_process_in_no}','999','{PROD_REQUEST_DATE}','{PROD_REQUEST_NO}','{G_CODE}', '','','{EMPL_NO}','{comboBox7.Text}01','OK',GETDATE(),'{EMPL_NO}',GETDATE(),'{EMPL_NO}','NM1')";

                                string next_process_lot_no = process_lot_no_generate(comboBox7.Text);

                                string insertvalueP501 = $"('002','{in_date}','{next_process_in_no}','001','001','{next_process_lot_no.Substring(5,3)}','','{next_process_lot_no}',GETDATE(),'{EMPL_NO}',GETDATE(),'{EMPL_NO}')";

                                pro.InsertYCSX(CTR_CD, PROD_REQUEST_DATE, PROD_REQUEST_NO, CODE_50, CODE_03, CODE_55, G_CODE, "A", PROD_REQUEST_QTY, CUST_CD, EMPL_NO, next_process_lot_no, EMPL_NO, EMPL_NO, DELIVERY_DT, PO_BALANCE.ToString(), TP.ToString(), BTP.ToString(), CK.ToString(), TOTAL_FCST.ToString(), W1.ToString(), W2.ToString(), W3.ToString(), W4.ToString(), W5.ToString(), W6.ToString(), W7.ToString(), W8.ToString(), PDuyet.ToString(),BLOCK_QTY.ToString());
                                pro.insertP500(insertvalueP500);
                                pro.insertP501(insertvalueP501);
                            }
                            else
                            {
                                pro.InsertYCSX(CTR_CD, PROD_REQUEST_DATE, PROD_REQUEST_NO, CODE_50, CODE_03, CODE_55, G_CODE, "A", PROD_REQUEST_QTY, CUST_CD, EMPL_NO, REMK, EMPL_NO, EMPL_NO, DELIVERY_DT, PO_BALANCE.ToString(), TP.ToString(), BTP.ToString(), CK.ToString(), TOTAL_FCST.ToString(), W1.ToString(), W2.ToString(), W3.ToString(), W4.ToString(), W5.ToString(), W6.ToString(), W7.ToString(), W8.ToString(), PDuyet.ToString(), BLOCK_QTY.ToString());
                            }
                            pro.writeHistory(CTR_CD, EMPL_NO, "YCSX TABLE", "THEM", "THEM YCSX", "0");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Loi");
                    }
                    MessageBox.Show("Đã hoàn thành thêm yêu cầu sản xuất");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }

        public string process_lot_no_generate(string machine_name)
        {

            //Machine name lấy khi trỏ vào 1 dòng nào đó trong bảng P500

            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            String sDate = DateTime.Now.ToString();
            DateTime datevalue = (Convert.ToDateTime(sDate.ToString()));
            int dy = datevalue.Day;
            int mn = datevalue.Month;
            int yy = datevalue.Year;
            string in_date = new Form1().STYMD(yy, mn, dy);
            //string in_date = STYMD(2022, 04, 27);
            // getlastest process_lot_no from machine name and in_date
            
            string NEXT_PROCESS_LOT_NO = machine_name + new Form1().CreateHeader2();

            dt = pro.getLastProcessLotNo(machine_name, in_date);
            if (dt.Rows.Count > 0)
            {
                // MessageBox.Show(dt.Rows[0]["PROCESS_LOT_NO"].ToString() + dt.Rows[0]["SEQ_NO"].ToString());
                NEXT_PROCESS_LOT_NO += String.Format("{0:000}", int.Parse(dt.Rows[0]["SEQ_NO"].ToString()) + 1);
            }
            else
            {
                // MessageBox.Show("Chưa có " + in_date);
                NEXT_PROCESS_LOT_NO += "001";
            }
            //MessageBox.Show(NEXT_PROCESS_LOT_NO);          
            return NEXT_PROCESS_LOT_NO;

        }


        private void button1_Click(object sender, EventArgs e)
        {
            addnewYCSX();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
