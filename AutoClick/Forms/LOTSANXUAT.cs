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
    public partial class LOTSANXUAT : Form
    {
        public LOTSANXUAT()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
           if(dataGridView1.Columns.Count >0)
            {
                dataGridView1.DataSource = null;
                dataGridView1.Columns.Remove("DANGKY");
                dataGridView1.Columns.Remove("SELECT");
            }
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dt = pro.getYCSXInfo2(textBox1.Text);
            if(dt.Rows.Count > 0)
            {
                label3.Text = "CODE: " +  dt.Rows[0]["G_NAME"].ToString();
                string G_CODE = dt.Rows[0]["G_CODE"].ToString();
                dt = pro.getFullBOMXuatLieu(G_CODE, "A");
                dataGridView1.DataSource = dt;

                DataGridViewCheckBoxColumn ck = new DataGridViewCheckBoxColumn();
                ck.Name = "SELECT";
                ck.HeaderText = "CHECK";
                ck.Width = 50;
                ck.ReadOnly = false;
                dataGridView1.Columns.Insert(0, ck);
                for (int r = 0; r < dataGridView1.Rows.Count; r++)
                {
                    dataGridView1.Rows[r].Cells["SELECT"].Value = false;
                }

                DataGridViewColumn dangky = new DataGridViewColumn();
                dangky.Name = "DANGKY";
                dangky.HeaderText = "METDANGKY";
                dangky.Width = 90;
                dangky.ReadOnly = false;
                dangky.CellTemplate = new DataGridViewTextBoxCell();
                dataGridView1.Columns.Insert(5, dangky);
                for (int r = 0; r < dataGridView1.Rows.Count; r++)
                {
                    dataGridView1.Rows[r].Cells["DANGKY"].Value = 0;
                }
            }
            else
            {
                MessageBox.Show("Không tòn tại số yêu cầu này");
            }

            label6.Text = "TÊN: " + pro.report_inspection_check_EMPLNAME(textBox2.Text);           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(dataGridView1.Rows.Count >0 && textBox1.Text != "" && textBox2.Text != "")
            {
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();
                dt = pro.getYCSXInfo2(textBox1.Text);
                if (dt.Rows.Count > 0)
                {
                    label3.Text = "CODE: " + dt.Rows[0]["G_NAME"].ToString();
                    string G_CODE = dt.Rows[0]["G_CODE"].ToString();
                    string CODE_50 = dt.Rows[0]["CODE_50"].ToString();
                    string PROD_REQUEST_DATE = dt.Rows[0]["PROD_REQUEST_DATE"].ToString();
                    string PROD_REQUEST_NO = dt.Rows[0]["PROD_REQUEST_NO"].ToString();
                    string OUT_DATE = "";
                    string OUT_NO = "";
                    string NEXT_OUT_NO = "";
                    string FACTORY = radioButton1.Checked == true ? "NM1" : "NM2";
                    string EMPL_NO = textBox2.Text;

                    //create OUT_DATE
                    String sDate = DateTime.Now.ToString();
                    DateTime datevalue = (Convert.ToDateTime(sDate.ToString()));
                    int dy = datevalue.Day;
                    int mn = datevalue.Month;
                    int yy = datevalue.Year;
                    OUT_DATE = new Form1().STYMD(yy, mn, dy);
                    //Generate NEXT_OUT_NO
                    dt = pro.getLastOutNo(OUT_DATE);
                    if (dt.Rows.Count > 0)
                    {
                        OUT_NO = dt.Rows[0]["OUT_NO"].ToString();
                        NEXT_OUT_NO = String.Format("{0:000}", int.Parse(OUT_NO) + 1);
                    }
                    else
                    {
                        NEXT_OUT_NO = "001";
                    }

                    //Insert Dang Ky Lieu O300
                    string insertValueO300 = $"('002', '{OUT_DATE}','{NEXT_OUT_NO}','01','{CODE_50}','01','{PROD_REQUEST_DATE}','{PROD_REQUEST_NO}','Y',GETDATE(), '{EMPL_NO}', '{FACTORY}')";
                    pro.InsertO300(insertValueO300);

                    //Insert Dang Ky Lieu O301

                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (!row.IsNewRow)
                        {
                            string
                                M_CODE = row.Cells["M_CODE"].Value.ToString(),
                                OUT_PRE_QTY = row.Cells["DANGKY"].Value.ToString();
                            bool SELECTED = Convert.ToBoolean(row.Cells["SELECT"].Value);
                            string NEXT_OUT_SEQ = "";
                                dt = pro.getLastOutSEQO301(OUT_DATE, NEXT_OUT_NO);

                            if(dt.Rows.Count >0)
                            {
                                string OUT_SEQ = dt.Rows[0]["OUT_SEQ"].ToString();
                                MessageBox.Show(OUT_SEQ);
                                NEXT_OUT_SEQ = String.Format("{0:000}", int.Parse(OUT_SEQ) + 1);
                            }
                            else
                            {
                                NEXT_OUT_SEQ = "001";
                            }
                            string insertValueO301 = $"('002','{OUT_DATE}','{NEXT_OUT_NO}','{NEXT_OUT_SEQ}', '01','{M_CODE}','{OUT_PRE_QTY}', 'Y', GETDATE(), '{EMPL_NO}')";

                            if(SELECTED == true)
                            {
                                pro.InsertO301(insertValueO301);
                                //MessageBox.Show($"{M_CODE}");
                            }                            
                        }
                    }



                }
                else
                {
                    MessageBox.Show("Không tòn tại số yêu cầu này");
                }

            }
            else
            {
                MessageBox.Show("Lỗi : Không có gì để đăng ký");
            }
           

        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void textBox3_MouseLeave(object sender, EventArgs e)
        {
           
                      

        }

        private void textBox4_MouseLeave(object sender, EventArgs e)
        {
           
        }

        private void textBox3_Leave(object sender, EventArgs e)
        {
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            string YCSX_NO = textBox3.Text;
            dt = pro.getYCSXInfo2(YCSX_NO);
            if (dt.Rows.Count > 0)
            {
                label9.Text = "CODE: " + dt.Rows[0]["G_NAME"].ToString();


            }
            else
            {
                label9.Text = "Không tòn tại số yêu cầu này";
            }
        }

        private void textBox4_Leave(object sender, EventArgs e)
        {
            ProductBLL pro = new ProductBLL();
            label10.Text = "TÊN: " + pro.report_inspection_check_EMPLNAME(textBox4.Text);
        }

        private void textBox8_Leave(object sender, EventArgs e)
        {

            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dt = pro.getYCSXInfo2(textBox8.Text);
            if (dt.Rows.Count > 0)
            {
                label17.Text = "CODE: " + dt.Rows[0]["G_NAME"].ToString();
            }
            else
            {
                label17.Text = "Không tòn tại số yêu cầu này";
            }

        }
        private void textBox7_Leave(object sender, EventArgs e)
        {
            ProductBLL pro = new ProductBLL();
            label16.Text = "TÊN: " + pro.report_inspection_check_EMPLNAME(textBox7.Text);
        }

        private void textBox5_Leave(object sender, EventArgs e)
        {
           
        }

        private void textBox6_Leave(object sender, EventArgs e)
        {
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dt = pro.check_M_NAME(textBox6.Text);

            if (dt.Rows.Count > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    label12.Text = row["M_NAME"].ToString();
                }
            }
            else
            {
                label12.Text = "KHÔNG TỒN TẠI LOT LIỆU NÀY";
            }            
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dt = pro.getYCSXInfo2(textBox1.Text);
            if (dt.Rows.Count > 0)
            {
                label3.Text = "CODE: " + dt.Rows[0]["G_NAME"].ToString();
            }
            else
            {
                label3.Text = "Không tòn tại số yêu cầu này";
            }
        }

        private void textBox2_Leave(object sender, EventArgs e)
        {
            ProductBLL pro = new ProductBLL();
            label6.Text = "TÊN: " + pro.report_inspection_check_EMPLNAME(textBox2.Text);
        }
    }
}
