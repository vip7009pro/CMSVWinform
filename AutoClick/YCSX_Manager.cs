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
    public partial class YCSX_Manager : Form
    {
        public YCSX_Manager()
        {
            InitializeComponent();
        }

       
        public void tratinhhinh(string codecms)
        {
            textBox2.Text = codecms;
            checkBox1.Checked = true;
            checkBox2.Checked = true;
            traYCSXManager();

        }
        private void YCSX_Manager_Load(object sender, EventArgs e)
        {
            comboBox1.Items.Add("Thong Thuong");
            comboBox1.Items.Add("SDI");
            comboBox1.Items.Add("GC");
            comboBox1.Items.Add("SAMPLE");
            comboBox1.Items.Add("NOT SAMPLE");
            comboBox1.Items.Add("ALL");
            comboBox1.Text = "ALL";
            this.ContextMenuStrip = contextMenuStrip1;
            this.dataGridView1.DefaultCellStyle.ForeColor = Color.Blue;
            this.dataGridView1.DefaultCellStyle.BackColor = Color.Beige;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.MultiSelect = true;
            dataGridView1.ReadOnly = true;

            if (!System.Windows.Forms.SystemInformation.TerminalServerSession)
            {
                Type dgvType = dataGridView1.GetType();
                PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                  BindingFlags.Instance | BindingFlags.NonPublic);
                pi.SetValue(dataGridView1, true, null);
            }

        }
        public string STYMD(int y, int m, int d)
        {
            string ymd, sty, stm, std;
            sty = "" + y;
            stm = "" + m;
            std = "" + d;
            if (m < 10)
            {
                stm = "0" + m;
            }
            if (d < 10)
            {
                std = "0" + d;
            }
            ymd = sty + stm + std;
            return ymd;
        }



        public string generate_condition_ycsxManager()
        {
            string query = "WHERE ";
            string fromdate, todate;
            fromdate = STYMD(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day);
            todate = STYMD(dateTimePicker2.Value.Year, dateTimePicker2.Value.Month, dateTimePicker2.Value.Day);
            //MessageBox.Show(fromdate);
            string ngaythang = " P400.PROD_REQUEST_DATE BETWEEN '" + fromdate + "' AND '" + todate + "' ";
            if (checkBox1.Checked == true)
            {
                ngaythang = "1=1 ";
            }

            string code;
            if (textBox3.Text != "")
            {
                code = "AND M100.G_NAME LIKE '%" + textBox3.Text + "%' ";
            }
            else
            {
                code = "";
            }

            string ycsxno;
            if (textBox4.Text != "")
            {
                ycsxno = "AND P400.PROD_REQUEST_NO= '" + textBox4.Text + "' ";
            }
            else
            {
                ycsxno = "";
            }


            string picname;
            if (textBox5.Text != "")
            {
                picname = "AND M010.EMPL_NAME LIKE '%" + textBox5.Text + "%' ";
            }
            else
            {
                picname = "";
            }

            string phanloai="";
            
            if (comboBox1.Text == "Thong Thuong")
            {
                phanloai = "AND P400.CODE_55= '01' ";
            }
            else if (comboBox1.Text == "SDI")
            {

                phanloai = "AND P400.CODE_55= '02' ";

            }
            else if (comboBox1.Text == "GC")
            {

                phanloai = "AND P400.CODE_55= '03' ";

            }
            else if (comboBox1.Text == "SAMPLE")
            {

                phanloai = "AND P400.CODE_55= '04' ";

            }
            else if (comboBox1.Text == "NOT SAMPLE")
            {

                phanloai = "AND P400.CODE_55<> '04' ";

            }
            else if (comboBox1.Text == "ALL")
            {

                phanloai = "";

            }
            else
            {
                phanloai = "AND P400.CODE_55= '999' ";
            }



            if (textBox5.Text != "")
            {
                picname = "AND M010.EMPL_NAME LIKE '%" + textBox5.Text + "%' ";
            }
            else
            {
                picname = "";
            }



            string cust_name_kd = "";
            if (textBox1.Text != "")
            {
                cust_name_kd = "AND M110.CUST_NAME_KD LIKE '%" + textBox1.Text + "%' ";
            }
            else
            {
                cust_name_kd = "";
            }

            
            string cmscode = "";
            if (textBox2.Text != "")
            {
                cmscode = "AND M100.G_CODE='" + textBox2.Text + "'";
            }
            else
            {
                cmscode = "";
            }

            string ycsx_pending = "";
            if(checkBox2.Checked)
            {
                ycsx_pending = " AND P400.YCSX_PENDING=1 ";
            }
            else
            {
                ycsx_pending = "";
            }

            string kieminput = "";
            if (checkBox3.Checked)
            {
                kieminput = " AND LOT_TOTAL_INPUT_QTY_EA<>0 "; 
            }
            else
            {
                kieminput = "";
            }

            string danhsachycsx = "";
            if(richTextBox1.Text!="")
            {
                danhsachycsx = "AND P400.PROD_REQUEST_NO IN (" + ycsx_list(richTextBox1) + ")";
            }
            else
            {
                danhsachycsx = "";
            }

            query += ngaythang + code + cust_name_kd + cmscode + ycsxno + picname + phanloai + ycsx_pending + kieminput + danhsachycsx;
            return query;
        }


        private void setRowNumber(DataGridView dgv)
        {
            foreach (DataGridViewRow row in dgv.Rows)
            {
                row.HeaderCell.Value = (row.Index + 1).ToString();
            }
        }
        public void changeHeaderText(DataGridView dtg1)
        {
            dtg1.Columns["LOT_TOTAL_INPUT_QTY_EA"].HeaderText = "NHẬP KIỂM";
            dtg1.Columns["LOT_TOTAL_OUTPUT_QTY_EA"].HeaderText = "XUẤT KIỂM";
            dtg1.Columns["INSPECT_BALANCE"].HeaderText = "TỒN KIỂM (CẢ LOSS)";
            dtg1.Columns["SHORTAGE_YCSX"].HeaderText = "TỒN YÊU CẦU";

            dtg1.Columns["G_CODE"].HeaderText = "CODE CMS";
            dtg1.Columns["G_NAME"].HeaderText = "CODE KHÁCH";
            dtg1.Columns["EMPL_NAME"].HeaderText = "NHÂN VIÊN KD";
            dtg1.Columns["CUST_NAME_KD"].HeaderText = "TÊN KHÁCH";
            dtg1.Columns["PROD_REQUEST_NO"].HeaderText = "SỐ YÊU CẦU";
            dtg1.Columns["PROD_REQUEST_DATE"].HeaderText = "NGÀY YÊU CẦU";
            dtg1.Columns["PROD_REQUEST_QTY"].HeaderText = "SL YÊU CẦU";


        }


        public void traYCSXManager()
        {
            try
            {
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();
                dt = pro.tra_YCSXMANAGER(generate_condition_ycsxManager());
                dataGridView1.DataSource = dt;
                setRowNumber(dataGridView1);
                formatYCSXTable(dataGridView1);
                changeHeaderText(dataGridView1);
                MessageBox.Show("Đã loát: " + dt.Rows.Count + " dòng");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi : " + ex.ToString());
            }

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
            dataGridView1.Columns["LOT_TOTAL_INPUT_QTY_EA"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["LOT_TOTAL_OUTPUT_QTY_EA"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["INSPECT_BALANCE"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["SHORTAGE_YCSX"].DefaultCellStyle.Format = "#,0";




            dataGridView1.Columns["LOT_TOTAL_INPUT_QTY_EA"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["LOT_TOTAL_INPUT_QTY_EA"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["LOT_TOTAL_INPUT_QTY_EA"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["LOT_TOTAL_OUTPUT_QTY_EA"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["LOT_TOTAL_OUTPUT_QTY_EA"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["LOT_TOTAL_OUTPUT_QTY_EA"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["INSPECT_BALANCE"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["INSPECT_BALANCE"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["INSPECT_BALANCE"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["SHORTAGE_YCSX"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["SHORTAGE_YCSX"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["SHORTAGE_YCSX"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);


            dataGridView1.Columns["G_NAME"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.BackColor = Color.Black;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["PROD_REQUEST_NO"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["PROD_REQUEST_NO"].DefaultCellStyle.BackColor = Color.Black;
            dataGridView1.Columns["PROD_REQUEST_NO"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

        }



        private void button1_Click(object sender, EventArgs e)
        {
            traYCSXManager();

        }

        private void sETHOÀNTHÀNHToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn thực sự muốn SET trạng thái YCSX là HOÀN THÀNH?", "SET STATUS ?", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {

                try
                {
                    ProductBLL pro = new ProductBLL();
                    DataTable dt = new DataTable();




                    var selectedRows = dataGridView1.SelectedRows
                          .OfType<DataGridViewRow>()
                          .Where(row => !row.IsNewRow)
                          .ToArray();
                    foreach (var row in selectedRows)
                    {
                        string ycsxno = row.Cells["PROD_REQUEST_NO"].Value.ToString();
                        //MessageBox.Show(ycsxno);
                        dt = pro.chang_YCSXMANAGER_STATUS(ycsxno, "0");
                    }


                    MessageBox.Show("Đã set hoàn thành cho " + dataGridView1.SelectedRows.Count + " dòng");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi : " + ex.ToString());
                }

            }

        }

        private void sETPENDINGToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn thực sự muốn SET trạng thái YCSX là PENDING?", "SET STATUS ?", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                try
                {
                    ProductBLL pro = new ProductBLL();
                    DataTable dt = new DataTable();




                    var selectedRows = dataGridView1.SelectedRows
                          .OfType<DataGridViewRow>()
                          .Where(row => !row.IsNewRow)
                          .ToArray();
                    foreach (var row in selectedRows)
                    {
                        string ycsxno = row.Cells["PROD_REQUEST_NO"].Value.ToString();
                        //MessageBox.Show(ycsxno);
                        dt = pro.chang_YCSXMANAGER_STATUS(ycsxno, "1");
                    }


                    MessageBox.Show("Đã set pending cho " + dataGridView1.SelectedRows.Count + " dòng");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi : " + ex.ToString());
                }

            }
        }

        private void tRALỊCHSỬXUẤTLIỆUToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();
                var selectedRows = dataGridView1.SelectedRows
                      .OfType<DataGridViewRow>()
                      .Where(row => !row.IsNewRow)
                      .ToArray();
                foreach (var row in selectedRows)
                {
                    string ycsxno = row.Cells["PROD_REQUEST_NO"].Value.ToString();
                    Material_History mh = new Material_History();
                    mh.ycsx_no = ycsxno;
                    mh.tra_Material_History();
                    mh.Show();
                }

                MessageBox.Show("Đã tra lịch sử xuất liệu cho yêu cầu được chọn");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi : " + ex.ToString());
            }
        }

        public string ycsx_list(RichTextBox richTextBox)
        {
            string ycsxlist = "'";
            for (int i = 0; i < richTextBox1.Lines.Length; i++)
            {
                ycsxlist += richTextBox1.Lines[i] + "','";
                //MessageBox.Show(richTextBox1.Lines[i]);
            }
            ycsxlist = ycsxlist.Substring(0,ycsxlist.Length -2);
            return ycsxlist;
            

        }
       
    }
}
