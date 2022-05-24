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
    public partial class InputForm : Form
    {
        public InputForm()
        {
            InitializeComponent();
        }

        public string LoginID = "";
        private void button1_Click(object sender, EventArgs e)
        {
        }

        private void InputForm_Load(object sender, EventArgs e)
        {
            comboBox1.Items.Add("PO");
            comboBox1.Items.Add("INVOICE");
            comboBox1.Items.Add("FCST");
            comboBox1.Items.Add("PLAN");
            comboBox1.Items.Add("CODE INFO");
            comboBox1.Items.Add("CUSTOMER INFO");
            if (!System.Windows.Forms.SystemInformation.TerminalServerSession)
            {
                Type dgvType = dataGridView1.GetType();
                PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                  BindingFlags.Instance | BindingFlags.NonPublic);
                pi.SetValue(dataGridView1, true, null);
            }

        }



        public void checkPO()
        {
            int total_flag = 0;
            if (total_flag == 0)
            {
                if (dataGridView1 == null || dataGridView1.Rows.Count == 0)
                {
                    MessageBox.Show("Import data vào trước  !");
                }
                else
                {
                    try
                    {
                        foreach (DataGridViewRow row in dataGridView1.Rows)
                        {
                            if (!row.IsNewRow)
                            {
                                string CUST_CD, EMPL_NO, G_CODE, PO_NO, PO_QTY, PO_DATE, RD_DATE, PROD_PRICE;
                                string CTR_CD = "002";

                                CUST_CD = Convert.ToString(row.Cells[0].Value);
                                EMPL_NO = Convert.ToString(row.Cells[1].Value);
                                G_CODE = Convert.ToString(row.Cells[2].Value);
                                PO_NO = Convert.ToString(row.Cells[3].Value);
                                PO_QTY = Convert.ToString(row.Cells[4].Value);
                                PO_DATE = Convert.ToString(row.Cells[5].Value);
                                RD_DATE = Convert.ToString(row.Cells[6].Value);
                                PROD_PRICE = Convert.ToString(row.Cells[7].Value);
                                //MessageBox.Show(CUST_CD + EMPL_NO + G_CODE + PO_NO + PO_QTY + PO_DATE + RD_DATE + PROD_PRICE);

                                if ((CUST_CD == "") || (EMPL_NO == "") || (G_CODE == "") || (PO_NO == "") || (PO_QTY == "") || (PO_DATE == "") || (RD_DATE == "") || (PROD_PRICE == ""))
                                {
                                    MessageBox.Show("Thông tin của sản phẩm: " + G_CODE + "có thông tin trống, sẽ ko thêm PO này !");
                                }
                                else
                                {
                                    ProductBLL pro = new ProductBLL();
                                    if (pro.checkPOExist(CUST_CD, G_CODE, PO_NO) == -1)
                                    {
                                        dataGridView1.Rows[row.Index].Cells[8].Value = "OK";
                                        dataGridView1.Rows[row.Index].Cells[8].Style.BackColor = Color.LightGreen;
                                    }
                                    else
                                    {
                                        dataGridView1.Rows[row.Index].Cells[8].Value = "NG - Đã tồn tại PO";
                                        dataGridView1.Rows[row.Index].Cells[8].Style.BackColor = Color.Red;
                                    }
                                }
                            }
                            else
                            {

                            }
                        }

                        MessageBox.Show("Đã hoàn thành check PO hàng loạt");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }

            }
            else
            {
                MessageBox.Show("Check choác cái gì ? Import vào r mới check được chứ ?!");
            }
        }



        public void upPOhangloat()
        {
            if (dataGridView1 == null || dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("Import data vào trước  !");
            }
            else
            {
                try
                {
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (!row.IsNewRow)
                        {
                            string CUST_CD, EMPL_NO, G_CODE, PO_NO, PO_QTY, PO_DATE, RD_DATE, PROD_PRICE;
                            string CTR_CD = "002";

                            CUST_CD = Convert.ToString(row.Cells[0].Value);
                            EMPL_NO = Convert.ToString(row.Cells[1].Value);
                            G_CODE = Convert.ToString(row.Cells[2].Value);
                            PO_NO = Convert.ToString(row.Cells[3].Value);
                            PO_QTY = Convert.ToString(row.Cells[4].Value);
                            PO_DATE = Convert.ToString(row.Cells[5].Value);
                            RD_DATE = Convert.ToString(row.Cells[6].Value);
                            PROD_PRICE = Convert.ToString(row.Cells[7].Value);


                            if ((CUST_CD == "") || (EMPL_NO == "") || (G_CODE == "") || (PO_NO == "") || (PO_QTY == "") || (PO_DATE == "") || (RD_DATE == "") || (PROD_PRICE == ""))
                            {
                                MessageBox.Show("Thông tin của sản phẩm: " + G_CODE + "có thông tin trống, sẽ ko thêm PO này !");
                            }
                            else
                            {
                                ProductBLL pro = new ProductBLL();

                                if (pro.checkPOExist(CUST_CD, G_CODE, PO_NO) == -1)
                                {
                                    pro.InsertPO(CTR_CD, CUST_CD, EMPL_NO, G_CODE, PO_NO, PO_QTY, PO_DATE, RD_DATE, PROD_PRICE);
                                    dataGridView1.Rows[row.Index].Cells[8].Value = "OK";
                                    dataGridView1.Rows[row.Index].Cells[8].Style.BackColor = Color.LightGreen;

                                }
                                else
                                {
                                    dataGridView1.Rows[row.Index].Cells[8].Value = "NG - Đã tồn tại PO";
                                    dataGridView1.Rows[row.Index].Cells[8].Style.BackColor = Color.Red;
                                }


                            }

                        }
                    }

                    MessageBox.Show("Đã hoàn thành thêm PO hàng loạt");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }

        public void checkInvoice()
        {
            int total_flag = 0;
            if (total_flag == 0)
            {
                if (dataGridView1 == null || dataGridView1.Rows.Count == 0)
                {
                    MessageBox.Show("Import data vào trước  !");
                }
                else
                {
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (!row.IsNewRow)
                        {

                            string CUST_CD, EMPL_NO, G_CODE, PO_NO, DELIVERY_QTY, DELIVERY_DATE, NOCANCEL;
                            string CTR_CD = "002";
                            try
                            {
                                CUST_CD = Convert.ToString(row.Cells[0].Value);
                                EMPL_NO = Convert.ToString(row.Cells[1].Value);
                                G_CODE = Convert.ToString(row.Cells[2].Value);
                                PO_NO = Convert.ToString(row.Cells[3].Value);
                                DELIVERY_QTY = Convert.ToString(row.Cells[4].Value);
                                DELIVERY_DATE = Convert.ToString(row.Cells[5].Value);
                                NOCANCEL = Convert.ToString(row.Cells[6].Value);

                                if ((CUST_CD == "") || (EMPL_NO == "") || (G_CODE == "") || (PO_NO == "") || (DELIVERY_QTY == "") || (DELIVERY_DATE == "") || (NOCANCEL == ""))
                                {
                                    MessageBox.Show("Thông tin của sản phẩm: " + G_CODE + "có thông tin trống, sẽ ko thêm PO này !");
                                }
                                else
                                {
                                    ProductBLL pro = new ProductBLL();
                                    int po_balance = pro.checkPOBalance(CUST_CD, G_CODE, PO_NO);
                                    if (int.Parse(DELIVERY_QTY) <= po_balance)
                                    {
                                        //pro.InsertInvoice(CTR_CD, CUST_CD, EMPL_NO, G_CODE, PO_NO, DELIVERY_QTY, DELIVERY_DATE, NOCANCEL);
                                        dataGridView1.Rows[row.Index].Cells[8].Value = "OK";
                                        dataGridView1.Rows[row.Index].Cells[8].Style.BackColor = Color.LightGreen;
                                    }
                                    else
                                    {
                                        dataGridView1.Rows[row.Index].Cells[8].Value = "NG - Giao hàng nhiều hơn PO";
                                        dataGridView1.Rows[row.Index].Cells[8].Style.BackColor = Color.Red;
                                    }

                                }
                            }
                            catch (Exception ex)
                            {
                                //MessageBox.Show(ex.ToString());   
                                //if(ex.InnerException is System.IndexOutOfRangeException)
                                // MessageBox.Show(ex.GetType().FullName.ToString());
                                if (ex.GetType().FullName == "System.IndexOutOfRangeException")
                                {
                                    dataGridView1.Rows[row.Index].Cells[7].Value = "NG-Không có PO";
                                    dataGridView1.Rows[row.Index].Cells[7].Style.BackColor = Color.Red;
                                }
                                else
                                {
                                    dataGridView1.Rows[row.Index].Cells[7].Style.BackColor = Color.Red;
                                    dataGridView1.Rows[row.Index].Cells[7].Value = "NG-" + ex.ToString();
                                }


                            }
                        }
                    }
                    MessageBox.Show("Đã hoàn thành check Invoice hàng loạt");

                }
            }
            else
            {
                MessageBox.Show("Check choác cái gì ! ?, import data từ file vào r mới check !");
            }
        }

        public void upInvoicehangloat()
        {
            if (dataGridView1 == null || dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("Import data vào trước  !");
            }
            else
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (!row.IsNewRow)
                    {

                        string CUST_CD, EMPL_NO, G_CODE, PO_NO, DELIVERY_QTY, DELIVERY_DATE, NOCANCEL;
                        string CTR_CD = "002";
                        try
                        {
                            CUST_CD = Convert.ToString(row.Cells[0].Value);
                            EMPL_NO = Convert.ToString(row.Cells[1].Value);
                            G_CODE = Convert.ToString(row.Cells[2].Value);
                            PO_NO = Convert.ToString(row.Cells[3].Value);
                            DELIVERY_QTY = Convert.ToString(row.Cells[4].Value);
                            DELIVERY_DATE = Convert.ToString(row.Cells[5].Value);
                            NOCANCEL = Convert.ToString(row.Cells[6].Value);



                            if ((CUST_CD == "") || (EMPL_NO == "") || (G_CODE == "") || (PO_NO == "") || (DELIVERY_QTY == "") || (DELIVERY_DATE == "") || (NOCANCEL == ""))
                            {
                                MessageBox.Show("Thông tin của sản phẩm: " + G_CODE + "có thông tin trống, sẽ ko thêm PO này !");
                            }
                            else
                            {
                                ProductBLL pro = new ProductBLL();
                                int po_balance = pro.checkPOBalance(CUST_CD, G_CODE, PO_NO);
                                if (int.Parse(DELIVERY_QTY) <= po_balance)
                                {
                                    pro.InsertInvoice(CTR_CD, CUST_CD, EMPL_NO, G_CODE, PO_NO, DELIVERY_QTY, DELIVERY_DATE, NOCANCEL);
                                    dataGridView1.Rows[row.Index].Cells[8].Value = "OK";
                                    dataGridView1.Rows[row.Index].Cells[8].Style.BackColor = Color.Green;
                                }
                                else
                                {
                                    dataGridView1.Rows[row.Index].Cells[8].Value = "NG - Giao hàng nhiều hơn PO";
                                    dataGridView1.Rows[row.Index].Cells[8].Style.BackColor = Color.Red;
                                }

                            }
                        }
                        catch (Exception ex)
                        {

                            if (ex.GetType().FullName == "System.IndexOutOfRangeException")
                            {
                                dataGridView1.Rows[row.Index].Cells[7].Value = "NG-Không có PO";
                                dataGridView1.Rows[row.Index].Cells[7].Style.BackColor = Color.Red;
                            }
                            else
                            {
                                dataGridView1.Rows[row.Index].Cells[7].Style.BackColor = Color.Red;
                                dataGridView1.Rows[row.Index].Cells[7].Value = "NG-" + ex.ToString();
                            }


                        }
                    }
                }
                MessageBox.Show("Đã hoàn thành thêm PO hàng loạt");



            }
        }

        public void updateINVOICENO()
        {
            if (dataGridView1 == null || dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("Import data vào trước  !");
            }
            else
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (!row.IsNewRow)
                    {

                        string CUST_CD, EMPL_NO, G_CODE, PO_NO, DELIVERY_QTY, DELIVERY_DATE, NOCANCEL, INVOICE_NO;
                        string CTR_CD = "002";
                        try
                        {
                            CUST_CD = Convert.ToString(row.Cells[0].Value);
                            EMPL_NO = Convert.ToString(row.Cells[1].Value);
                            G_CODE = Convert.ToString(row.Cells[2].Value);
                            PO_NO = Convert.ToString(row.Cells[3].Value);
                            DELIVERY_QTY = Convert.ToString(row.Cells[4].Value);
                            DELIVERY_DATE = Convert.ToString(row.Cells[5].Value);
                            NOCANCEL = Convert.ToString(row.Cells[6].Value);
                            INVOICE_NO = Convert.ToString(row.Cells[7].Value);


                            if ((CUST_CD == "") || (EMPL_NO == "") || (G_CODE == "") || (PO_NO == "") || (DELIVERY_QTY == "") || (DELIVERY_DATE == "") || (NOCANCEL == ""))
                            {
                                MessageBox.Show("Thông tin của sản phẩm: " + G_CODE + "có thông tin trống, sẽ ko thêm INVOICE này !");
                            }
                            else
                            {
                                ProductBLL pro = new ProductBLL();
                                int po_balance = pro.checkPOBalance(CUST_CD, G_CODE, PO_NO);
                                if (po_balance >= 0)
                                {
                                    pro.updateINVOICE_NO(CTR_CD, CUST_CD, EMPL_NO, G_CODE, PO_NO, INVOICE_NO);
                                    dataGridView1.Rows[row.Index].Cells[9].Value = "OK";
                                    dataGridView1.Rows[row.Index].Cells[9].Style.BackColor = Color.Green;
                                }
                                else
                                {
                                    dataGridView1.Rows[row.Index].Cells[9].Value = "NG - Giao hàng nhiều hơn PO";
                                    dataGridView1.Rows[row.Index].Cells[9].Style.BackColor = Color.Red;
                                }

                            }
                        }
                        catch (Exception ex)
                        {

                            if (ex.GetType().FullName == "System.IndexOutOfRangeException")
                            {
                                dataGridView1.Rows[row.Index].Cells[7].Value = "NG-Không có PO";
                                dataGridView1.Rows[row.Index].Cells[7].Style.BackColor = Color.Red;
                            }
                            else
                            {
                                dataGridView1.Rows[row.Index].Cells[7].Style.BackColor = Color.Red;
                                dataGridView1.Rows[row.Index].Cells[7].Value = "NG-" + ex.ToString();
                            }


                        }
                    }
                }
                MessageBox.Show("Đã hoàn thành thêm Invoice hàng loạt");

            }
        }


        public void insertFCST(int check)
        {
            if (check == 1)
            {
                if (dataGridView1 == null || dataGridView1.Rows.Count == 0)
                {
                    MessageBox.Show("Import data vào trước  !");
                }
                else
                {
                    try
                    {
                        foreach (DataGridViewRow row in dataGridView1.Rows)
                        {
                            if (!row.IsNewRow)
                            {
                                string CUST_CD, EMPL_NO, G_CODE, YEAR, WEEKNO, PROD_PRICE, W1, W2, W3, W4, W5, W6, W7, W8, W9, W10, W11, W12, W13, W14, W15, W16, W17, W18, W19, W20, W21, W22;
                                string CTR_CD = "002";

                                EMPL_NO = Convert.ToString(row.Cells[0].Value);
                                CUST_CD = Convert.ToString(row.Cells[1].Value);
                                G_CODE = Convert.ToString(row.Cells[2].Value);
                                PROD_PRICE = Convert.ToString(row.Cells[3].Value);
                                YEAR = Convert.ToString(row.Cells[4].Value);
                                WEEKNO = Convert.ToString(row.Cells[5].Value);
                                W1 = Convert.ToString(row.Cells[6].Value);
                                W2 = Convert.ToString(row.Cells[7].Value);
                                W3 = Convert.ToString(row.Cells[8].Value);
                                W4 = Convert.ToString(row.Cells[9].Value);
                                W5 = Convert.ToString(row.Cells[10].Value);
                                W6 = Convert.ToString(row.Cells[11].Value);
                                W7 = Convert.ToString(row.Cells[12].Value);
                                W8 = Convert.ToString(row.Cells[13].Value);
                                W9 = Convert.ToString(row.Cells[14].Value);
                                W10 = Convert.ToString(row.Cells[15].Value);
                                W11 = Convert.ToString(row.Cells[16].Value);
                                W12 = Convert.ToString(row.Cells[17].Value);
                                W13 = Convert.ToString(row.Cells[18].Value);
                                W14 = Convert.ToString(row.Cells[19].Value);
                                W15 = Convert.ToString(row.Cells[20].Value);
                                W16 = Convert.ToString(row.Cells[21].Value);
                                W17 = Convert.ToString(row.Cells[22].Value);
                                W18 = Convert.ToString(row.Cells[23].Value);
                                W19 = Convert.ToString(row.Cells[24].Value);
                                W20 = Convert.ToString(row.Cells[25].Value);
                                W21 = Convert.ToString(row.Cells[26].Value);
                                W22 = Convert.ToString(row.Cells[27].Value);
                                if (W1.IndexOf("-", 0) == 0 || W2.IndexOf("-", 0) == 0 || W3.IndexOf("-", 0) == 0 || W4.IndexOf("-", 0) == 0 || W5.IndexOf("-", 0) == 0 || W6.IndexOf("-", 0) == 0 || W7.IndexOf("-", 0) == 0 || W8.IndexOf("-", 0) == 0 || W9.IndexOf("-", 0) == 0 || W10.IndexOf("-", 0) == 0 || W11.IndexOf("-", 0) == 0 || W12.IndexOf("-", 0) == 0 || W13.IndexOf("-", 0) == 0 || W14.IndexOf("-", 0) == 0 || W15.IndexOf("-", 0) == 0 || W16.IndexOf("-", 0) == 0 || W17.IndexOf("-", 0) == 0 || W18.IndexOf("-", 0) == 0 || W19.IndexOf("-", 0) == 0 || W20.IndexOf("-", 0) == 0 || W21.IndexOf("-", 0) == 0 || W22.IndexOf("-", 0) == 0)
                                {
                                    MessageBox.Show("Không được phép có giá trị âm, bỏ qua dòng này");
                                    dataGridView1.Rows[row.Index].Cells[29].Value = "NG - Giá trị fcst âm";
                                    dataGridView1.Rows[row.Index].Cells[29].Style.BackColor = Color.Red;
                                }
                                else if ((EMPL_NO == "") || (CUST_CD == "") || (G_CODE == "") || (PROD_PRICE == "") || (YEAR == "") || (WEEKNO == "") || (W1 == "") || (W2 == "") || (W3 == "") || (W4 == "") || (W5 == "") || (W6 == "") || (W7 == "") || (W8 == "") || (W9 == "") || (W10 == "") || (W11 == "") || (W12 == "") || (W13 == "") || (W14 == "") || (W15 == "") || (W16 == "") || (W17 == "") || (W18 == "") || (W19 == "") || (W20 == "") || (W21 == "") || (W22 == ""))
                                {
                                    MessageBox.Show("Thông tin của sản phẩm: " + G_CODE + "có thông tin trống, sẽ ko thêm PO này !");
                                }
                                else
                                {
                                    ProductBLL pro = new ProductBLL();
                                    if (pro.checkFCSTExist(CUST_CD, G_CODE, YEAR, WEEKNO) == -1)
                                    {
                                        pro.InsertFCST(CTR_CD, EMPL_NO, CUST_CD, G_CODE, PROD_PRICE, YEAR, WEEKNO, W1, W2, W3, W4, W5, W6, W7, W8, W9, W10, W11, W12, W13, W14, W15, W16, W17, W18, W19, W20, W21, W22);
                                        dataGridView1.Rows[row.Index].Cells[29].Value = "OK";
                                        dataGridView1.Rows[row.Index].Cells[29].Style.BackColor = Color.LightGreen;
                                    }
                                    else
                                    {
                                        dataGridView1.Rows[row.Index].Cells[29].Value = "NG - Đã tồn tại FCST";
                                        dataGridView1.Rows[row.Index].Cells[29].Style.BackColor = Color.Red;
                                    }


                                }

                            }
                        }

                        MessageBox.Show("Đã hoàn thành check FCST hàng loạt");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }

            }
            else
            {
                if (dataGridView1 == null || dataGridView1.Rows.Count == 0)
                {
                    MessageBox.Show("Import data vào trước  !");
                }
                else
                {
                    try
                    {
                        foreach (DataGridViewRow row in dataGridView1.Rows)
                        {
                            if (!row.IsNewRow)
                            {
                                string CUST_CD, EMPL_NO, G_CODE, YEAR, WEEKNO, PROD_PRICE, W1, W2, W3, W4, W5, W6, W7, W8, W9, W10, W11, W12, W13, W14, W15, W16, W17, W18, W19, W20, W21, W22;
                                string CTR_CD = "002";

                                EMPL_NO = Convert.ToString(row.Cells[0].Value);
                                CUST_CD = Convert.ToString(row.Cells[1].Value);
                                G_CODE = Convert.ToString(row.Cells[2].Value);
                                PROD_PRICE = Convert.ToString(row.Cells[3].Value);
                                YEAR = Convert.ToString(row.Cells[4].Value);
                                WEEKNO = Convert.ToString(row.Cells[5].Value);
                                W1 = Convert.ToString(row.Cells[6].Value);
                                W2 = Convert.ToString(row.Cells[7].Value);
                                W3 = Convert.ToString(row.Cells[8].Value);
                                W4 = Convert.ToString(row.Cells[9].Value);
                                W5 = Convert.ToString(row.Cells[10].Value);
                                W6 = Convert.ToString(row.Cells[11].Value);
                                W7 = Convert.ToString(row.Cells[12].Value);
                                W8 = Convert.ToString(row.Cells[13].Value);
                                W9 = Convert.ToString(row.Cells[14].Value);
                                W10 = Convert.ToString(row.Cells[15].Value);
                                W11 = Convert.ToString(row.Cells[16].Value);
                                W12 = Convert.ToString(row.Cells[17].Value);
                                W13 = Convert.ToString(row.Cells[18].Value);
                                W14 = Convert.ToString(row.Cells[19].Value);
                                W15 = Convert.ToString(row.Cells[20].Value);
                                W16 = Convert.ToString(row.Cells[21].Value);
                                W17 = Convert.ToString(row.Cells[22].Value);
                                W18 = Convert.ToString(row.Cells[23].Value);
                                W19 = Convert.ToString(row.Cells[24].Value);
                                W20 = Convert.ToString(row.Cells[25].Value);
                                W21 = Convert.ToString(row.Cells[26].Value);
                                W22 = Convert.ToString(row.Cells[27].Value);


                                if (W1.IndexOf("-", 0) == 0 || W2.IndexOf("-", 0) == 0 || W3.IndexOf("-", 0) == 0 || W4.IndexOf("-", 0) == 0 || W5.IndexOf("-", 0) == 0 || W6.IndexOf("-", 0) == 0 || W7.IndexOf("-", 0) == 0 || W8.IndexOf("-", 0) == 0 || W9.IndexOf("-", 0) == 0 || W10.IndexOf("-", 0) == 0 || W11.IndexOf("-", 0) == 0 || W12.IndexOf("-", 0) == 0 || W13.IndexOf("-", 0) == 0 || W14.IndexOf("-", 0) == 0 || W15.IndexOf("-", 0) == 0 || W16.IndexOf("-", 0) == 0 || W17.IndexOf("-", 0) == 0 || W18.IndexOf("-", 0) == 0 || W19.IndexOf("-", 0) == 0 || W20.IndexOf("-", 0) == 0 || W21.IndexOf("-", 0) == 0 || W22.IndexOf("-", 0) == 0)
                                {
                                    MessageBox.Show("Không được phép có giá trị âm, bỏ qua dòng này");
                                    dataGridView1.Rows[row.Index].Cells[29].Value = "NG - Giá trị fcst âm";
                                    dataGridView1.Rows[row.Index].Cells[29].Style.BackColor = Color.Red;
                                }
                                else if((EMPL_NO == "") || (CUST_CD == "") || (G_CODE == "") || (PROD_PRICE == "") || (YEAR == "") || (WEEKNO == "") || (W1 == "") || (W2 == "") || (W3 == "") || (W4 == "") || (W5 == "") || (W6 == "") || (W7 == "") || (W8 == "") || (W9 == "") || (W10 == "") || (W11 == "") || (W12 == "") || (W13 == "") || (W14 == "") || (W15 == "") || (W16 == "") || (W17 == "") || (W18 == "") || (W19 == "") || (W20 == "") || (W21 == "") || (W22 == ""))
                                {
                                    MessageBox.Show("Thông tin của sản phẩm: " + G_CODE + "có thông tin trống, sẽ ko thêm PO này !");
                                }
                                else
                                {
                                    ProductBLL pro = new ProductBLL();
                                    if (pro.checkFCSTExist(CUST_CD, G_CODE, YEAR, WEEKNO) == -1)
                                    {
                                        // pro.InsertFCST(CTR_CD, EMPL_NO, CUST_CD, G_CODE, PROD_PRICE, YEAR, WEEKNO, W1, W2, W3, W4, W5, W6, W7, W8, W9, W10, W11, W12, W13, W14, W15, W16, W17, W18, W19, W20, W21, W22);
                                        dataGridView1.Rows[row.Index].Cells[29].Value = "OK";
                                        dataGridView1.Rows[row.Index].Cells[29].Style.BackColor = Color.LightGreen;
                                    }
                                    else
                                    {
                                        dataGridView1.Rows[row.Index].Cells[29].Value = "NG - Đã tồn tại FCST";
                                        dataGridView1.Rows[row.Index].Cells[29].Style.BackColor = Color.Red;
                                    }


                                }

                            }
                        }

                        MessageBox.Show("Đã hoàn thành thêm PO hàng loạt");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }

            }

        }

        public void insertKHGH(int check)
        {
            if (check == 1)
            {
                if (dataGridView1 == null || dataGridView1.Rows.Count == 0)
                {
                    MessageBox.Show("Import data vào trước  !");
                }
                else
                {
                    try
                    {
                        foreach (DataGridViewRow row in dataGridView1.Rows)
                        {
                            if (!row.IsNewRow)
                            {
                                string CUST_CD, EMPL_NO, G_CODE, PLAN_DATE, D1, D2, D3, D4, D5, D6, D7, D8, REMARK;
                                string CTR_CD = "002";

                                EMPL_NO = Convert.ToString(row.Cells[0].Value);
                                CUST_CD = Convert.ToString(row.Cells[1].Value);
                                G_CODE = Convert.ToString(row.Cells[2].Value);
                                PLAN_DATE = Convert.ToString(row.Cells[3].Value);
                                D1 = Convert.ToString(row.Cells[4].Value);
                                D2 = Convert.ToString(row.Cells[5].Value);
                                D3 = Convert.ToString(row.Cells[6].Value);
                                D4 = Convert.ToString(row.Cells[7].Value);
                                D5 = Convert.ToString(row.Cells[8].Value);
                                D6 = Convert.ToString(row.Cells[9].Value);
                                D7 = Convert.ToString(row.Cells[10].Value);
                                D8 = Convert.ToString(row.Cells[11].Value);
                                REMARK = Convert.ToString(row.Cells[12].Value);


                                if (D1.IndexOf("-", 0) == 0 || D2.IndexOf("-", 0) == 0 || D3.IndexOf("-", 0) == 0 || D4.IndexOf("-", 0) == 0 || D5.IndexOf("-", 0) == 0 || D6.IndexOf("-", 0) == 0 || D7.IndexOf("-", 0) == 0 || D8.IndexOf("-", 0) == 0)
                                {
                                    MessageBox.Show("Không được phép có giá trị âm, bỏ qua dòng này");
                                    dataGridView1.Rows[row.Index].Cells[29].Value = "NG - Giá trị plan âm";
                                    dataGridView1.Rows[row.Index].Cells[29].Style.BackColor = Color.Red;
                                }
                                else if ((EMPL_NO == "") || (CUST_CD == "") || (G_CODE == "") || (PLAN_DATE == "") || (D1 == "") || (D2 == "") || (D3 == "") || (D4 == "") || (D5 == "") || (D6 == "") || (D7 == "") || (D8 == ""))
                                {
                                    MessageBox.Show("Thông tin của sản phẩm: " + G_CODE + "có thông tin trống, sẽ ko thêm KHGH này !");
                                }
                                else
                                {
                                    ProductBLL pro = new ProductBLL();
                                    if (pro.checkKHGHExist(CUST_CD, G_CODE, PLAN_DATE) == -1)
                                    {
                                        pro.InsertPlan(CTR_CD, EMPL_NO, CUST_CD, G_CODE, PLAN_DATE, D1, D2, D3, D4, D5, D6, D7, D8, REMARK);
                                        pro.writeHistory(CTR_CD, LoginID, "PLAN TABLE", "THEM", "THEM PLAN GIAO HANG", "0");
                                        dataGridView1.Rows[row.Index].Cells[13].Value = "OK";
                                        dataGridView1.Rows[row.Index].Cells[13].Style.BackColor = Color.LightGreen;
                                    }
                                    else
                                    {
                                        dataGridView1.Rows[row.Index].Cells[13].Value = "NG - Đã tồn tại FCST";
                                        dataGridView1.Rows[row.Index].Cells[13].Style.BackColor = Color.Red;
                                    }


                                }

                            }
                        }

                        MessageBox.Show("Đã hoàn thành check PLAN hàng loạt");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }

            }
            else
            {
                if (dataGridView1 == null || dataGridView1.Rows.Count == 0)
                {
                    MessageBox.Show("Import data vào trước  !");
                }
                else
                {
                    try
                    {
                        foreach (DataGridViewRow row in dataGridView1.Rows)
                        {
                            if (!row.IsNewRow)
                            {
                                string CUST_CD, EMPL_NO, G_CODE, PLAN_DATE, D1, D2, D3, D4, D5, D6, D7, D8, REMARK;
                                string CTR_CD = "002";

                                EMPL_NO = Convert.ToString(row.Cells[0].Value);
                                CUST_CD = Convert.ToString(row.Cells[1].Value);
                                G_CODE = Convert.ToString(row.Cells[2].Value);
                                PLAN_DATE = Convert.ToString(row.Cells[3].Value);
                                D1 = Convert.ToString(row.Cells[4].Value);
                                D2 = Convert.ToString(row.Cells[5].Value);
                                D3 = Convert.ToString(row.Cells[6].Value);
                                D4 = Convert.ToString(row.Cells[7].Value);
                                D5 = Convert.ToString(row.Cells[8].Value);
                                D6 = Convert.ToString(row.Cells[9].Value);
                                D7 = Convert.ToString(row.Cells[10].Value);
                                D8 = Convert.ToString(row.Cells[11].Value);
                                REMARK = Convert.ToString(row.Cells[12].Value);


                                if (D1.IndexOf("-", 0) == 0 || D2.IndexOf("-", 0) == 0 || D3.IndexOf("-", 0) == 0 || D4.IndexOf("-", 0) == 0 || D5.IndexOf("-", 0) == 0 || D6.IndexOf("-", 0) == 0 || D7.IndexOf("-", 0) == 0 || D8.IndexOf("-", 0) == 0)
                                {
                                    MessageBox.Show("Không được phép có giá trị âm, bỏ qua dòng này");
                                    dataGridView1.Rows[row.Index].Cells[29].Value = "NG - Giá trị plan âm";
                                    dataGridView1.Rows[row.Index].Cells[29].Style.BackColor = Color.Red;
                                }
                                else if ((EMPL_NO == "") || (CUST_CD == "") || (G_CODE == "") || (PLAN_DATE == "") || (D1 == "") || (D2 == "") || (D3 == "") || (D4 == "") || (D5 == "") || (D6 == "") || (D7 == "") || (D8 == ""))
                                {
                                    MessageBox.Show("Thông tin của sản phẩm: " + G_CODE + "có thông tin trống, sẽ ko thêm KHGH này !");
                                }
                                else
                                {
                                    ProductBLL pro = new ProductBLL();
                                    if (pro.checkKHGHExist(CUST_CD, G_CODE, PLAN_DATE) == -1)
                                    {
                                        //pro.InsertPlan(CTR_CD, EMPL_NO, CUST_CD, G_CODE, PLAN_DATE, D1, D2, D3, D4, D5, D6, D7, D8);
                                        dataGridView1.Rows[row.Index].Cells[13].Value = "OK";
                                        dataGridView1.Rows[row.Index].Cells[13].Style.BackColor = Color.LightGreen;
                                    }
                                    else
                                    {
                                        dataGridView1.Rows[row.Index].Cells[13].Value = "NG - Đã tồn tại PLAN";
                                        dataGridView1.Rows[row.Index].Cells[13].Style.BackColor = Color.Red;
                                    }


                                }

                            }
                        }

                        MessageBox.Show("Đã hoàn thành thêm PLAN hàng loạt");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }

            }

        }


        public void updatecodeinfo()
        {
            if (dataGridView1 == null || dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("Import data vào trước  !");
            }
            else
            {
                try
                {
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (!row.IsNewRow)
                        {
                            string G_CODE, G_CODE_KD, PROD_TYPE, PROD_MODEL, PROD_PROJECT, PROD_MAIN_MATERIAL;

                            G_CODE = Convert.ToString(row.Cells[0].Value);
                            G_CODE_KD = Convert.ToString(row.Cells[2].Value);
                            PROD_TYPE = Convert.ToString(row.Cells[3].Value);
                            PROD_MODEL = Convert.ToString(row.Cells[5].Value);
                            PROD_PROJECT = Convert.ToString(row.Cells[6].Value);
                            PROD_MAIN_MATERIAL = Convert.ToString(row.Cells[4].Value);

                            ProductBLL pro = new ProductBLL();
                            DataTable dt = new DataTable();
                            pro.updateInfo(G_CODE, G_CODE_KD, PROD_TYPE, PROD_MODEL, PROD_PROJECT, PROD_MAIN_MATERIAL);

                        }
                    }

                    MessageBox.Show("Đã hoàn thành update Code info hàng loạt");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }

        }


        public void uploadcustomerinfor()
        {
            if (dataGridView1 == null || dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("Import data vào trước  !");
            }
            else
            {
                try
                {
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (!row.IsNewRow)
                        {
                            string CUST_CD, CUST_NAME_KD;

                            CUST_CD = Convert.ToString(row.Cells[0].Value);
                            CUST_NAME_KD = Convert.ToString(row.Cells[2].Value);

                            ProductBLL pro = new ProductBLL();
                            DataTable dt = new DataTable();
                            pro.updateCustomer(CUST_CD, CUST_NAME_KD);


                        }
                    }

                    MessageBox.Show("Đã hoàn thành update thông tin khách hàng hàng loạt");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //MessageBox.Show("Chuyển form sang : " + comboBox1.Text);
            if (comboBox1.Text == "PO")
            {
                label2.Text = " NHẬP PO";
                dataGridView1.DataSource = null;
                dataGridView1.ColumnCount = 11;
                dataGridView1.Columns[0].Name = "Mã Khách Hàng";
                dataGridView1.Columns[1].Name = "Mã Nhân Viên";
                dataGridView1.Columns[2].Name = "Code CMS";
                dataGridView1.Columns[3].Name = "PO No";
                dataGridView1.Columns[4].Name = "PO QTY";
                dataGridView1.Columns[5].Name = "PO DATE";
                dataGridView1.Columns[6].Name = "RD DATE";
                dataGridView1.Columns[7].Name = "PRICE";
                dataGridView1.Columns[8].Name = "EXT1";
                dataGridView1.Columns[9].Name = "EXT2";
                dataGridView1.Columns[10].Name = "EXT3";



            }
            else if (comboBox1.Text == "INVOICE")
            {
                label2.Text = " NHẬP INVOICE";
                dataGridView1.DataSource = null;
                dataGridView1.ColumnCount = 11;
                dataGridView1.Columns[0].Name = "Mã Khách Hàng";
                dataGridView1.Columns[1].Name = "Mã Nhân Viên";
                dataGridView1.Columns[2].Name = "Code CMS";
                dataGridView1.Columns[3].Name = "PO No";
                dataGridView1.Columns[4].Name = "DELIVERY QTY";
                dataGridView1.Columns[5].Name = "DELIVERY DATE";
                dataGridView1.Columns[6].Name = "NO CANCEL";
                dataGridView1.Columns[7].Name = "INVOICE NO";
                dataGridView1.Columns[8].Name = "EXT1";
                dataGridView1.Columns[9].Name = "EXT2";
                dataGridView1.Columns[10].Name = "EXT3";
            }
            else if (comboBox1.Text == "FCST")
            {
                label2.Text = " NHẬP FCST";
                dataGridView1.DataSource = null;
                dataGridView1.ColumnCount = 31;
                dataGridView1.Columns[0].Name = "Mã Nhân Viên";
                dataGridView1.Columns[1].Name = "Mã Khách Hàng";
                dataGridView1.Columns[2].Name = "Code CMS";
                dataGridView1.Columns[3].Name = "PRICE";
                dataGridView1.Columns[4].Name = "YEAR";
                dataGridView1.Columns[5].Name = "WEEK";
                dataGridView1.Columns[6].Name = "W+1";
                dataGridView1.Columns[7].Name = "W+2";
                dataGridView1.Columns[8].Name = "W+3";
                dataGridView1.Columns[9].Name = "W+4";
                dataGridView1.Columns[10].Name = "W+5";
                dataGridView1.Columns[11].Name = "W+6";
                dataGridView1.Columns[12].Name = "W+7";
                dataGridView1.Columns[13].Name = "W+8";
                dataGridView1.Columns[14].Name = "W+9";
                dataGridView1.Columns[15].Name = "W+10";
                dataGridView1.Columns[16].Name = "W+11";
                dataGridView1.Columns[17].Name = "W+12";
                dataGridView1.Columns[18].Name = "W+13";
                dataGridView1.Columns[19].Name = "W+14";
                dataGridView1.Columns[20].Name = "W+15";
                dataGridView1.Columns[21].Name = "W+16";
                dataGridView1.Columns[22].Name = "W+17";
                dataGridView1.Columns[23].Name = "W+18";
                dataGridView1.Columns[24].Name = "W+19";
                dataGridView1.Columns[25].Name = "W+20";
                dataGridView1.Columns[26].Name = "W+21";
                dataGridView1.Columns[27].Name = "W+22";
                dataGridView1.Columns[28].Name = "EXT1";
                dataGridView1.Columns[29].Name = "EXT2";
                dataGridView1.Columns[30].Name = "EXT3";


            }
            else if (comboBox1.Text == "PLAN")
            {
                label2.Text = " NHẬP PLAN";
                dataGridView1.DataSource = null;
                dataGridView1.ColumnCount = 16;
                dataGridView1.Columns[0].Name = "Mã Nhân Viên";
                dataGridView1.Columns[1].Name = "Mã Khách Hàng";
                dataGridView1.Columns[2].Name = "Code CMS";
                dataGridView1.Columns[3].Name = "PLAN DATE";
                dataGridView1.Columns[4].Name = "D+0";
                dataGridView1.Columns[5].Name = "D+1";
                dataGridView1.Columns[6].Name = "D+2";
                dataGridView1.Columns[7].Name = "D+3";
                dataGridView1.Columns[8].Name = "D+4";
                dataGridView1.Columns[9].Name = "D+5";
                dataGridView1.Columns[10].Name = "D+6";
                dataGridView1.Columns[11].Name = "D+7";
                dataGridView1.Columns[12].Name = "REMARK";
                dataGridView1.Columns[13].Name = "EXT1";
                dataGridView1.Columns[14].Name = "EXT2";
                dataGridView1.Columns[15].Name = "EXT3";

            }
            
            else if (comboBox1.Text == "CODE INFO")
            {
                label2.Text = " NHẬP THÔNG TIN CODE";
                dataGridView1.DataSource = null;
                dataGridView1.ColumnCount = 0;
                /*
                dataGridView1.ColumnCount = 16;
                dataGridView1.Columns[0].Name = "CODE CMS";
                dataGridView1.Columns[1].Name = "CODE Kinh Doanh";
                dataGridView1.Columns[2].Name = "Phân Loại";
                dataGridView1.Columns[3].Name = "MODEL";
                dataGridView1.Columns[4].Name = "PROJECT";
                dataGridView1.Columns[5].Name = "MATERIAL";

                dataGridView1.Columns[6].Name = "EXT1";
                dataGridView1.Columns[7].Name = "EXT2";
                dataGridView1.Columns[8].Name = "EXT3";
                */
            }
            else if (comboBox1.Text == "CUSTOMER INFO")
            {
                label2.Text = " NHẬP THÔNG TIN KHÁCH HÀNG";
                dataGridView1.DataSource = null;                
                dataGridView1.ColumnCount = 0;
                /*
                dataGridView1.Columns[0].Name = "MÃ KHÁCH HÀNG";
                dataGridView1.Columns[1].Name = "TÊN KHÁCH HÀNG  (KD)";
                dataGridView1.Columns[2].Name = "EXT1";
                dataGridView1.Columns[3].Name = "EXT2";
                dataGridView1.Columns[4].Name = "EXT3";
                */
            }
            

        }

        private void checkPOToolStripMenuItem_Click(object sender, EventArgs e)
        {
            checkPO();
        }

        private void uploadHàngLoạtToolStripMenuItem_Click(object sender, EventArgs e)
        {
            upPOhangloat();
        }

        private void checkInvoiceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            checkInvoice();
        }

        private void uploadHàngLoạtToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            upInvoicehangloat();
        }

        private void updateINVOICENOToolStripMenuItem_Click(object sender, EventArgs e)
        {
            updateINVOICENO();
        }

        private void checkFCSTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            insertFCST(0);

        }

        private void uploadHàngLoạtToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            insertFCST(1);
        }

        private void checkPLANToolStripMenuItem_Click(object sender, EventArgs e)
        {
            insertKHGH(0);
        }

        private void uploadHàngLoạtToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            insertKHGH(1);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dt = pro.getnameandcode(textBox1.Text);
            comboBox2.DataSource = dt;
            comboBox2.DisplayMember = "G_NAME";
            comboBox2.ValueMember = "G_NAME";
            comboBox3.DataSource = dt;
            comboBox3.DisplayMember = "G_CODE";
            comboBox3.ValueMember = "G_CODE";
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dt = pro.getcustomerinfo(textBox2.Text);
            comboBox5.DataSource = dt;
            comboBox5.DisplayMember = "CUST_NAME";
            comboBox5.ValueMember = "CUST_NAME";
            comboBox4.DataSource = dt;
            comboBox4.DisplayMember = "CUST_CD";
            comboBox4.ValueMember = "CUST_CD";
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            try
            {
                if (comboBox1.Text == "PO")
                {
                    int rowindex = dataGridView1.CurrentRow.Index;
                    dataGridView1.Rows[rowindex].Cells[0].Value = comboBox4.Text;
                    dataGridView1.Rows[rowindex].Cells[1].Value = LoginID;
                    dataGridView1.Rows[rowindex].Cells[2].Value = comboBox3.Text;

                }
                else if (comboBox1.Text == "INVOICE")
                {
                    int rowindex = dataGridView1.CurrentRow.Index;
                    dataGridView1.Rows[rowindex].Cells[0].Value = comboBox4.Text;
                    dataGridView1.Rows[rowindex].Cells[1].Value = LoginID;
                    dataGridView1.Rows[rowindex].Cells[2].Value = comboBox3.Text;
                }
                else if (comboBox1.Text == "FCST")
                {
                    int rowindex = dataGridView1.CurrentRow.Index;
                    dataGridView1.Rows[rowindex].Cells[0].Value = LoginID;
                    dataGridView1.Rows[rowindex].Cells[1].Value = comboBox4.Text;
                    dataGridView1.Rows[rowindex].Cells[2].Value = comboBox3.Text;
                }
                else if (comboBox1.Text == "PLAN")
                {
                    int rowindex = dataGridView1.CurrentRow.Index;
                    dataGridView1.Rows[rowindex].Cells[0].Value = LoginID;
                    dataGridView1.Rows[rowindex].Cells[1].Value = comboBox4.Text;
                    dataGridView1.Rows[rowindex].Cells[2].Value = comboBox3.Text;
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi : " + ex.ToString());
            }

        }

        private void upCodeInfoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if(textBox1.Text =="update")
            {
                updatecodeinfo();
            }
            else
            {
                MessageBox.Show("Muốn update nhập update vào ô tên code");
            }
            
        }

        private void upCustomerInfoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "update")
            {
                uploadcustomerinfor();
            }
            else
            {
                MessageBox.Show("Muốn update nhập update vào ô tên code");
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text == "CODE INFO")
            {
                this.dataGridView1.DataSource = null;
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();
                dt = pro.report_NoData();
                this.dataGridView1.DataSource = null;
                this.dataGridView1.Rows.Clear();
                this.dataGridView1.DataSource = dt;
            }
            else if (comboBox1.Text == "CUSTOMER INFO")
            {
                this.dataGridView1.DataSource = null;
                ProductBLL pro = new ProductBLL();
                DataTable dt = pro.report_getCustomerList();
                this.dataGridView1.DataSource = null;
                this.dataGridView1.Rows.Clear();
                this.dataGridView1.DataSource = dt;
            }
        }
        private void testDtgvToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }
    }
}
