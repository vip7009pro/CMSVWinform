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
    public partial class NewYCSX : Form
    {
        public string Login_ID = "NHU1903";
        public NewYCSX()
        {
            InitializeComponent();
        }

        private void NewYCSX_Load(object sender, EventArgs e)
        {
            if (!System.Windows.Forms.SystemInformation.TerminalServerSession)
            {
                Type dgvType = dataGridView1.GetType();
                PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                  BindingFlags.Instance | BindingFlags.NonPublic);
                pi.SetValue(dataGridView1, true, null);
            }
            if (!System.Windows.Forms.SystemInformation.TerminalServerSession)
            {
                Type dgvType = dataGridView2.GetType();
                PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                  BindingFlags.Instance | BindingFlags.NonPublic);
                pi.SetValue(dataGridView2, true, null);
            }
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
            this.dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
            this.dataGridView2.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;


            ProductBLL pro = new ProductBLL();
            DataTable dt = pro.report_getCustomerList();

            comboBoxCustomer.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            comboBoxCustomer.AutoCompleteSource = AutoCompleteSource.ListItems;
            comboBoxCustomer.DataSource = dt;
            comboBoxCustomer.ValueMember = "CUST_CD";
            comboBoxCustomer.DisplayMember = "CUST_NAME_KD";
            dt = pro.info_getCODEInfo("");
            comboBoxCode.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            comboBoxCode.AutoCompleteSource = AutoCompleteSource.ListItems;
            comboBoxCode.DataSource = dt;
            comboBoxCode.ValueMember = "G_CODE";
            comboBoxCode.DisplayMember = "G_NAME";

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

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (dataGridView1 == null || dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("Import data vào trước  !");
            }
            else
            {
                ProductBLL pro1 = new ProductBLL();
                String ngaygiohethong = pro1.getsystemDateTime();
                String ngayhethong = ngaygiohethong.Substring(0, 4) + ngaygiohethong.Substring(5, 2) + ngaygiohethong.Substring(8, 2);
                try
                {
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (!row.IsNewRow)
                        {
                            string PROD_REQUEST_DATE, CODE_50, CODE_55, G_CODE, RIV_NO, PROD_REQUEST_QTY, CUST_CD, EMPL_NO, REMK, DELIVERY_DT, NGUYENCHIEC;
                            string CTR_CD = "002";
                            string CODE_03 = "01";
                            int TP = 0, BTP = 0, CK = 0, BLOCK_QTY =0,  W1 = 0, W2 = 0, W3 = 0, W4 = 0, W5 = 0, W6 = 0, W7 = 0, W8 = 0, PO_BALANCE = 0, TOTAL_FCST = 0, PDuyet = 0;
                            PROD_REQUEST_DATE = ngayhethong;
                            CODE_55 = Convert.ToString(row.Cells["LOAI_SAN_XUAT"].Value);
                            CODE_50 = Convert.ToString(row.Cells["LOAI_XUAT_HANG"].Value);                           
                            G_CODE = Convert.ToString(row.Cells["G_CODE"].Value);
                            RIV_NO = "A";
                            PROD_REQUEST_QTY = Convert.ToString(row.Cells["PROD_REQUEST_QTY"].Value);
                            CUST_CD = Convert.ToString(row.Cells["CUST_CD"].Value);
                            EMPL_NO = Convert.ToString(row.Cells["EMPL_NO"].Value);
                            REMK = Convert.ToString(row.Cells["REMK"].Value);
                            DELIVERY_DT = Convert.ToString(row.Cells["DELIVERY_DT"].Value);
                            NGUYENCHIEC = Convert.ToString(row.Cells["NGUYEN_CHIEC"].Value);

                            if ((PROD_REQUEST_DATE == "") || (CODE_50 == "") || (CODE_55 == "") || (G_CODE == "") || (RIV_NO == "") || (PROD_REQUEST_QTY == "") || (CUST_CD == "") || (EMPL_NO == "") || (DELIVERY_DT == ""))
                            {
                                MessageBox.Show("Yêu cầu của sản phẩm: " + G_CODE + "có thông tin trống, sẽ ko yêu cầu trường hợp này !");
                            }
                            else
                            {
                                ProductBLL pro = new ProductBLL();
                                DataTable dt = new DataTable();
                                dt = pro.getLastYCSXNo();
                                string lastycsxno = "";
                                if (dt.Rows.Count > 0)
                                {
                                    int yccuoiint = 0;
                                    lastycsxno = dt.Rows[0]["PROD_REQUEST_NO"].ToString();
                                    if (lastycsxno.Substring(0, 3) !=  new Form1().CreateHeader2())
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

                                    string PROD_REQUEST_NO = new Form1().CreateHeader2() + lastycsxno;

                                    int check_riv = pro.checkRIV_NO(G_CODE, RIV_NO);

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
                                        dataGridView1.Rows[row.Index].Cells["G_CODE"].Style.BackColor = Color.Red;
                                    }
                                    else
                                    {
                                        if (tonkho_gcode.Rows.Count > 0)
                                        {
                                            TP = int.Parse(tonkho_gcode.Rows[0]["TON_TP"].ToString());
                                            BTP = int.Parse(tonkho_gcode.Rows[0]["BTP"].ToString());
                                            CK = int.Parse(tonkho_gcode.Rows[0]["TONG_TON_KIEM"].ToString());
                                            BLOCK_QTY = int.Parse(tonkho_gcode.Rows[0]["BLOCK_QTY"].ToString());
                                        }
                                        if (fcst_gcode.Rows.Count > 0)
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
                                        if (CODE_55 == "04") PDuyet = 1;
                                        if (pobalance_gcode.Rows.Count > 0)
                                        {
                                            PO_BALANCE = int.Parse(pobalance_gcode.Rows[0]["PO_BALANCE"].ToString());
                                            if (PO_BALANCE > 0) PDuyet = 1;
                                        }

                                        if (NGUYENCHIEC!="")
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

                                            string next_process_lot_no = new Form3().process_lot_no_generate(NGUYENCHIEC);

                                            string insertvalueP501 = $"('002','{in_date}','{next_process_in_no}','001','001','{next_process_lot_no.Substring(5, 3)}','','{next_process_lot_no}',GETDATE(),'{EMPL_NO}',GETDATE(),'{EMPL_NO}')";

                                            switch (CODE_55)
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
                                                    CODE_55 = "01";
                                                    break;

                                            }
                                            switch (CODE_50)
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
                                                    CODE_50 = "07";
                                                    break;
                                            }


                                            pro.InsertYCSX(CTR_CD, PROD_REQUEST_DATE, PROD_REQUEST_NO, CODE_50, CODE_03, CODE_55, G_CODE, "A", PROD_REQUEST_QTY, CUST_CD, EMPL_NO, next_process_lot_no, EMPL_NO, EMPL_NO, DELIVERY_DT, PO_BALANCE.ToString(), TP.ToString(), BTP.ToString(), CK.ToString(), TOTAL_FCST.ToString(), W1.ToString(), W2.ToString(), W3.ToString(), W4.ToString(), W5.ToString(), W6.ToString(), W7.ToString(), W8.ToString(), PDuyet.ToString(), BLOCK_QTY.ToString());
                                            pro.insertP500(insertvalueP500);
                                            pro.insertP501(insertvalueP501);
                                        }
                                        else
                                        {
                                            switch (CODE_55)
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
                                            switch (CODE_50)
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
                                            pro.InsertYCSX(CTR_CD, PROD_REQUEST_DATE, PROD_REQUEST_NO, CODE_50, CODE_03, CODE_55, G_CODE, "A", PROD_REQUEST_QTY, CUST_CD, EMPL_NO, REMK, EMPL_NO, EMPL_NO, DELIVERY_DT, PO_BALANCE.ToString(), TP.ToString(), BTP.ToString(), CK.ToString(), TOTAL_FCST.ToString(), W1.ToString(), W2.ToString(), W3.ToString(), W4.ToString(), W5.ToString(), W6.ToString(), W7.ToString(), W8.ToString(), PDuyet.ToString(), BLOCK_QTY.ToString());
                                        }

                                    }

                                }
                                else
                                {
                                    MessageBox.Show("Loi");
                                }
                            }

                        }
                    }

                    MessageBox.Show("Đã hoàn thành yêu cầu sản xuất hàng loạt");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }

        private void comboBoxCode_DisplayMemberChanged(object sender, EventArgs e)
        {
           
        }

        private void comboBoxCode_SelectedValueChanged(object sender, EventArgs e)
        {
            label7.Text = comboBoxCode.SelectedValue.ToString();
        }

        public List<string> listchuabanve = null;
        List<YeuCauSanXuat> dsNV = null;

        private void button3_Click(object sender, EventArgs e)
        {                   

            try
            {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.ShowDialog();
                string file = ofd.FileName;

                if (file != "")
                {
                    dataGridView1.Show();
                    dsNV = ExcelFactory.readFromExcelFile(file);
                    this.dataGridView1.DataSource = null;
                    this.dataGridView1.Columns.Clear();
                    dataGridView1.DataSource = dsNV;                    
                   

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            if (dataGridView1 == null || dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("Import data vào trước  !");
            }
            else
            {
                ProductBLL pro1 = new ProductBLL();
                String ngaygiohethong = pro1.getsystemDateTime();
                String ngayhethong = ngaygiohethong.Substring(0, 4) + ngaygiohethong.Substring(5, 2) + ngaygiohethong.Substring(8, 2);

                try
                {
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (!row.IsNewRow)
                        {
                            string PROD_REQUEST_DATE, CODE_50, CODE_55, G_CODE, RIV_NO, PROD_REQUEST_QTY, CUST_CD, EMPL_NO, REMK, DELIVERY_DT;
                            string CTR_CD = "002";
                            string CODE_03 = "01";
                            int TP = 0, BTP = 0, CK = 0, BLOCK_QTY=0, W1 = 0, W2 = 0, W3 = 0, W4 = 0, W5 = 0, W6 = 0, W7 = 0, W8 = 0, PO_BALANCE = 0, TOTAL_FCST = 0, PDuyet = 0;
                            PROD_REQUEST_DATE = ngayhethong;
                            CODE_50 = Convert.ToString(row.Cells[1].Value);
                            CODE_55 = Convert.ToString(row.Cells[2].Value);
                            G_CODE = Convert.ToString(row.Cells[3].Value);
                            RIV_NO = Convert.ToString(row.Cells[4].Value);
                            PROD_REQUEST_QTY = Convert.ToString(row.Cells[5].Value);
                            CUST_CD = Convert.ToString(row.Cells[6].Value);
                            EMPL_NO = Convert.ToString(row.Cells[7].Value);
                            REMK = Convert.ToString(row.Cells[8].Value);
                            DELIVERY_DT = Convert.ToString(row.Cells[9].Value);


                            if ((PROD_REQUEST_DATE == "") || (CODE_50 == "") || (CODE_55 == "") || (G_CODE == "") || (RIV_NO == "") || (PROD_REQUEST_QTY == "") || (CUST_CD == "") || (EMPL_NO == "") || (DELIVERY_DT == ""))
                            {
                                MessageBox.Show("Yêu cầu của sản phẩm: " + G_CODE + "có thông tin trống, sẽ ko yêu cầu trường hợp này !");
                            }
                            else
                            {
                                ProductBLL pro = new ProductBLL();
                                DataTable dt = new DataTable();
                                dt = pro.getLastYCSXNo();
                                string lastycsxno = "";
                                if (dt.Rows.Count > 0)
                                {
                                    int yccuoiint = 0;
                                    lastycsxno = dt.Rows[0]["PROD_REQUEST_NO"].ToString();
                                    if (lastycsxno.Substring(0, 3) != new Form1().CreateHeader2())
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

                                    string PROD_REQUEST_NO = new Form1().CreateHeader2() + lastycsxno;

                                    int check_riv = pro.checkRIV_NO(G_CODE, RIV_NO);

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
                                        if (tonkho_gcode.Rows.Count > 0)
                                        {
                                            TP = int.Parse(tonkho_gcode.Rows[0]["TON_TP"].ToString());
                                            BTP = int.Parse(tonkho_gcode.Rows[0]["BTP"].ToString());
                                            CK = int.Parse(tonkho_gcode.Rows[0]["TONG_TON_KIEM"].ToString());
                                            BLOCK_QTY = int.Parse(tonkho_gcode.Rows[0]["BLOCK_QTY"].ToString());
                                        }
                                        if (fcst_gcode.Rows.Count > 0)
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
                                        if (CODE_55 == "04") PDuyet = 1;
                                        if (pobalance_gcode.Rows.Count > 0)
                                        {
                                            PO_BALANCE = int.Parse(pobalance_gcode.Rows[0]["PO_BALANCE"].ToString());
                                            if (PO_BALANCE > 0 ) PDuyet = 1;
                                        }

                                        pro.InsertYCSX(CTR_CD, PROD_REQUEST_DATE, PROD_REQUEST_NO, CODE_50, CODE_03, CODE_55, G_CODE, "A", PROD_REQUEST_QTY, CUST_CD, EMPL_NO, REMK, EMPL_NO, EMPL_NO, DELIVERY_DT, PO_BALANCE.ToString(), TP.ToString(), BTP.ToString(), CK.ToString(), TOTAL_FCST.ToString(), W1.ToString(), W2.ToString(), W3.ToString(), W4.ToString(), W5.ToString(), W6.ToString(), W7.ToString(), W8.ToString(), PDuyet.ToString(), BLOCK_QTY.ToString());
                                        pro.writeHistory(CTR_CD, EMPL_NO, "YCSX TABLE", "THEM", "THEM YCSX", "0");
                                    }

                                }
                                else
                                {
                                    MessageBox.Show("Loi");
                                }
                            }


                        }
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
            string PROD_REQUEST_DATE = "", CODE_50 = "", CODE_55 = "", G_CODE = "", G_NAME = "", PROD_REQUEST_QTY = "", CUST_NAME_KD = "", CUST_CD = "", PIC_KD = "", REMARK = "", DELIVERY_DT = "", NGUYENCHIEC = "";

            ProductBLL pro = new ProductBLL();
            String ngaygiohethong = pro.getsystemDateTime();
            String ngayhethong = ngaygiohethong.Substring(0, 4) + ngaygiohethong.Substring(5, 2) + ngaygiohethong.Substring(8, 2);
            PROD_REQUEST_DATE = ngayhethong;
            CODE_50 = comboBox5.Text;
            CODE_55 = comboBox6.Text;
            G_CODE = label7.Text;
            G_NAME = comboBoxCode.Text;
            CUST_NAME_KD = comboBoxCustomer.Text;
            CUST_CD = comboBoxCustomer.SelectedValue.ToString();
            PIC_KD = Login_ID;
            REMARK = textBox4.Text;
            DELIVERY_DT = textBox5.Text;
            PROD_REQUEST_QTY = textBox3.Text;

            if (checkBox1.Checked == true)
            {
                NGUYENCHIEC = comboBox7.Text;
            }

            if (CODE_50 == "")
            {
                MessageBox.Show("Không để trống phân loại sản xuất");
            }
            else if (CODE_55 == "")
            {
                MessageBox.Show("Không để trống phân loại giao hàng");
            }
            else if (PROD_REQUEST_QTY == "")
            {
                MessageBox.Show("Không để trống số lượng YCSX");
            }
            else if (DELIVERY_DT == "")
            {
                MessageBox.Show("Không để trống ngày giao hàng");
            }
            else if (checkBox1.Checked == true && comboBox7.Text == "")
            {
                MessageBox.Show("Hãy lựa chọn kiểu hàng nguyên chiếc");
            }
            else
            {
                this.dataGridView1.Rows.Add(PROD_REQUEST_DATE, CODE_50, CODE_55, G_CODE, G_NAME, PROD_REQUEST_QTY, CUST_NAME_KD, CUST_CD, PIC_KD, REMARK, DELIVERY_DT, NGUYENCHIEC);
            }


        }

        private void textBox5_KeyDown(object sender, KeyEventArgs e)
        {
           
        }


        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                ProductBLL pro = new ProductBLL();
                DataTable dt = pro.traYCSXCODEKD(textBox1.Text);
                dataGridView2.DataSource = dt;
                new Form1().setRowNumber(dataGridView2);
                formatYCSXTable(dataGridView2);
            }
        }

        private void dataGridView2_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewRow row = dataGridView2.Rows[e.RowIndex];
            comboBoxCode.SelectedValue = row.Cells["G_CODE"].Value.ToString();
            comboBoxCustomer.SelectedValue = row.Cells["CUST_CD"].Value.ToString();
           
            //comboBox1.SelectedItem = row.Cells["PROD_TYPE"].Value.ToString();
        }

        public void formatYCSXTable(DataGridView dataGridView1)
        {
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;

            dataGridView1.Columns["PROD_REQUEST_QTY"].DefaultCellStyle.Format = "#,0";      


            dataGridView1.Columns["G_NAME"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.BackColor = Color.Black;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["PROD_REQUEST_NO"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["PROD_REQUEST_NO"].DefaultCellStyle.BackColor = Color.Black;
            dataGridView1.Columns["PROD_REQUEST_NO"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

        }

    }
}
