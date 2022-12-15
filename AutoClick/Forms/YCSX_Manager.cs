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
using System.IO;
using RawPrint;
using Spire;

namespace AutoClick
{
    public partial class YCSX_Manager : Form
    {
        public string Login_ID = "";
        public YCSX_Manager()
        {
            InitializeComponent();
        }

        public DataTable dtgv1_data = new DataTable();
        public int my_ycsx_flag = 0;
        public int all_ycsx_flag = 0;
        public int xuatfile_print_ycsx_flag = 0;
        public int xuatfile_only_flag = 0;
        public int xuatchithi_flag = 0;
        public int checkbanve_flag = 0;
        public int set_hoanthanh_flag = 0;
        public int set_pending_flag = 0;
        public int xoa_ycsx_flag = 0;
        public int pheduyet_ycsx_flag = 0;

        public string saveFolderPath = "";
        public void tratinhhinh(string codecms)
        {
            textBox2.Text = codecms;
            checkBox1.Checked = true;
            checkBox2.Checked = true;
            traYCSXManager();
        }


        private void YCSX_Manager_Load(object sender, EventArgs e)
        {
            pictureBox1.Hide();
            int h = Screen.PrimaryScreen.WorkingArea.Height;
            int w = Screen.PrimaryScreen.WorkingArea.Width;
            this.ClientSize = new Size(w, h);

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
            dataGridView1.MultiSelect = true;
            dataGridView1.ReadOnly = false;

            if (!System.Windows.Forms.SystemInformation.TerminalServerSession)
            {
                Type dgvType = dataGridView1.GetType();
                PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                  BindingFlags.Instance | BindingFlags.NonPublic);
                pi.SetValue(dataGridView1, true, null);
            }
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
            this.dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

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
                if (dataGridView1.Columns.Count > 0)
                {
                    dataGridView1.DataSource = null;
                    dataGridView1.Columns.Remove("SELECT");
                    //dataGridView1.Columns.Remove("button");
                }

                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();
                dtgv1_data = pro.tra_YCSXMANAGER(generate_condition_ycsxManager());

                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi : " + ex.ToString());
            }

        }


        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            var senderGrid = (DataGridView)sender;

            if (senderGrid.Columns[e.ColumnIndex] is DataGridViewButtonColumn &&
                e.RowIndex >= 0)
            {
                MessageBox.Show("Index: " + e.RowIndex);
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

            dataGridView1.Columns["PO_TDYCSX"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["TOTAL_TKHO_TDYCSX"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["TKHO_TDYCSX"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["BTP_TDYCSX"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["CK_TDYCSX"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["BLOCK_TDYCSX"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["FCST_TDYCSX"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W1"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W2"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W3"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W4"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W5"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W6"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W7"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W8"].DefaultCellStyle.Format = "#,0";
            




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



            dataGridView1.Columns["PO_TDYCSX"].DefaultCellStyle.ForeColor = Color.Blue;
            dataGridView1.Columns["PO_TDYCSX"].DefaultCellStyle.BackColor = Color.LightGray;
            dataGridView1.Columns["PO_TDYCSX"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Regular);


            dataGridView1.Columns["TOTAL_TKHO_TDYCSX"].DefaultCellStyle.ForeColor = Color.Blue;
            dataGridView1.Columns["TOTAL_TKHO_TDYCSX"].DefaultCellStyle.BackColor = Color.LightGray;
            dataGridView1.Columns["TOTAL_TKHO_TDYCSX"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Regular);

            dataGridView1.Columns["FCST_TDYCSX"].DefaultCellStyle.ForeColor = Color.Blue;
            dataGridView1.Columns["FCST_TDYCSX"].DefaultCellStyle.BackColor = Color.LightGray;
            dataGridView1.Columns["FCST_TDYCSX"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Regular);
        }



        private void button1_Click(object sender, EventArgs e)
        {
            my_ycsx_flag = 0;
            all_ycsx_flag = 1;
            xuatfile_print_ycsx_flag = 0;
            xuatfile_only_flag = 0;
            xuatchithi_flag = 0;
            checkbanve_flag = 0;
            set_hoanthanh_flag = 0;
            set_pending_flag = 0;
            xoa_ycsx_flag = 0;
            pheduyet_ycsx_flag = 0;

            if (!backgroundWorker1.IsBusy)
            {
                pictureBox1.Show();
                backgroundWorker1.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Tiến trình khác đang chạy, thử lại sau");
            }

            //traYCSXManager();

        }

        public void sethoanthanh()
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
                    progressBar1.Maximum = selectedRows.Length;
                    int ii = 0;
                    foreach (var row in selectedRows)
                    {
                        string ycsxno = row.Cells["PROD_REQUEST_NO"].Value.ToString();
                        //MessageBox.Show(ycsxno);
                        dt = pro.chang_YCSXMANAGER_STATUS(ycsxno, "0");
                        ii++;
                        backgroundWorker1.ReportProgress(ii);
                    }
                    MessageBox.Show("Đã set hoàn thành cho " + dataGridView1.SelectedRows.Count + " dòng");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi : " + ex.ToString());
                }
            }
        }
        private void sETHOÀNTHÀNHToolStripMenuItem_Click(object sender, EventArgs e)
        {
            my_ycsx_flag = 0;
            all_ycsx_flag = 0;
            xuatfile_print_ycsx_flag = 0;
            xuatfile_only_flag = 0;
            xuatchithi_flag = 0;
            checkbanve_flag = 0;
            set_hoanthanh_flag = 1;
            set_pending_flag = 0;
            xoa_ycsx_flag = 0;
            pheduyet_ycsx_flag = 0;

            if (!backgroundWorker1.IsBusy)
            {
                pictureBox1.Show();
                backgroundWorker1.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Tiến trình khác đang chạy, thử lại sau");
            }

            /*
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
            }*/
        }

        public void setpending()
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
                    progressBar1.Maximum = selectedRows.Length;
                    int ii = 0;
                    foreach (var row in selectedRows)
                    {
                        string ycsxno = row.Cells["PROD_REQUEST_NO"].Value.ToString();
                        //MessageBox.Show(ycsxno);
                        dt = pro.chang_YCSXMANAGER_STATUS(ycsxno, "1");
                        ii++;
                        backgroundWorker1.ReportProgress(ii);
                    }
                    MessageBox.Show("Đã set pending cho " + dataGridView1.SelectedRows.Count + " dòng");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi : " + ex.ToString());
                }

            }
        }
        private void sETPENDINGToolStripMenuItem_Click(object sender, EventArgs e)
        {
            my_ycsx_flag = 0;
            all_ycsx_flag = 0;
            xuatfile_print_ycsx_flag = 0;
            xuatfile_only_flag = 0;
            xuatchithi_flag = 0;
            checkbanve_flag = 0;
            set_hoanthanh_flag = 0;
            set_pending_flag = 1;
            xoa_ycsx_flag = 0;
            pheduyet_ycsx_flag = 0;

            if (!backgroundWorker1.IsBusy)
            {
                pictureBox1.Show();
                backgroundWorker1.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Tiến trình khác đang chạy, thử lại sau");
            }

            /*
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
            */
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

        private void checkBảnVẽToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void newYCSXToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 frm3 = new Form3();
            frm3.loginIDfrm3 = Login_ID;
            frm3.Show();
        }

        private void thêmYêuCầuMớiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void thêm1YCToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 frm3 = new Form3();
            frm3.loginIDfrm3 = Login_ID;
            frm3.Show();
        }

        private void thêmNhiềuYCSXToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            NewYCSX newycsx = new NewYCSX();
            newycsx.Login_ID = Login_ID;
            newycsx.Show();
        }

        private void thêmNhiềuYCSXToolStripMenuItem_Click(object sender, EventArgs e)
        {
            NewYCSX newycsx = new NewYCSX();
            newycsx.Login_ID = Login_ID;
            newycsx.Show();
        }
        public List<string> listchuabanve = null;
        public void xuatfile_inycsx(bool inhaykhong)
        {
            CheckBox cb1 = new CheckBox();
            cb1.Checked = inhaykhong;
            dataGridView1.EndEdit();
            try
            {
                listchuabanve = null;
                string Dir = System.IO.Directory.GetCurrentDirectory();
                //MessageBox.Show(Dir);
                string file = Dir + "\\template.xlsx";
                string saveycsxpath = "";


                using (var fbd = new FolderBrowserDialog())
                {
                    DialogResult result = fbd.ShowDialog();

                    if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                    {
                        saveycsxpath = fbd.SelectedPath;

                        ProductBLL pro = new ProductBLL();
                        DataTable dt = new DataTable();

                        //dataGridView1.Rows[rowindex].Cells[columnindex].Value.ToString();
                        int checkRowsCount = 0;
                        foreach (DataGridViewRow row in dataGridView1.Rows)
                        { 
                            if (!row.IsNewRow)
                            {
                                if (row.Cells["SELECT"].Value.ToString() == "True")
                                {
                                    checkRowsCount++;
                                }
                            }
                        }
                        progressBar1.Minimum = 0; //Đặt giá trị nhỏ nhất cho ProgressBar
                        progressBar1.Maximum = checkRowsCount; //Đặt giá trị lớn nhất cho ProgressBar
                        

                        int startprogress = 0;

                        foreach (DataGridViewRow row in dataGridView1.Rows)
                        {
                            if (!row.IsNewRow)
                            {
                                if (row.Cells["SELECT"].Value.ToString() == "True")
                                {
                                    
                                    string ycsxno = row.Cells["PROD_REQUEST_NO"].Value.ToString();
                                    dt = pro.getFullInfo(ycsxno);
                                    if (file != "")
                                    {
                                        string drawfilename = dt.Rows[0]["G_NAME"].ToString().Substring(0, 11) + "_" + dt.Rows[0]["G_CODE"].ToString().Substring(7, 1) + ".pdf";
                                        string pdffile = Dir + "\\BANVE\\" + drawfilename;
                                        ExcelFactory.editFileExcel(file, dt, cb1, saveycsxpath);
                                        if(inhaykhong == true)
                                        {
                                            if (File.Exists(pdffile))
                                            {
                                                new Form1().printPDF(pdffile);
                                            }
                                            else
                                            {
                                                MessageBox.Show("Không có bản vẽ : " + dt.Rows[0]["G_NAME"].ToString());
                                            }

                                        }
                                        startprogress = startprogress + 1;
                                        backgroundWorker1.ReportProgress(startprogress);
                                        //progressBar1.Value = startprogress;

                                    }
                                    
                                }

                            }
                                                       
                        }
                        MessageBox.Show("Export Yêu cầu hoàn thành !");
                        progressBar1.Value = 0;
                        // MessageBox.Show(saveycsxpath);
                    }
                }



            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.ToString());
            }
        }


        public void xuatfile_inycsx2(bool inhaykhong)
        {
            CheckBox cb1 = new CheckBox();
            cb1.Checked = inhaykhong;
            dataGridView1.EndEdit();
            try
            {
                listchuabanve = null;
                string Dir = System.IO.Directory.GetCurrentDirectory();
                //MessageBox.Show(Dir);
                string file = Dir + "\\template.xlsx";
                string saveycsxpath = ""; 
                if (saveFolderPath != "")
                {
                    saveycsxpath = saveFolderPath;
                    //MessageBox.Show(saveFolderPath);
                    ProductBLL pro = new ProductBLL();
                    DataTable dt = new DataTable();

                    //dataGridView1.Rows[rowindex].Cells[columnindex].Value.ToString();
                    int checkRowsCount = 0;
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (!row.IsNewRow)
                        {
                            if (row.Cells["SELECT"].Value.ToString() == "True")
                            {
                                checkRowsCount++;
                            }
                        }
                    }
                    progressBar1.Minimum = 0; //Đặt giá trị nhỏ nhất cho ProgressBar
                    progressBar1.Maximum = checkRowsCount; //Đặt giá trị lớn nhất cho ProgressBar


                    int startprogress = 0;

                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (!row.IsNewRow)
                        {
                            if (row.Cells["SELECT"].Value.ToString() == "True")
                            {
                                string ycsxno = row.Cells["PROD_REQUEST_NO"].Value.ToString();
                                string PDUYET = row.Cells["PDUYET"].Value.ToString();
                                if(PDUYET == "1")
                                {
                                    dt = pro.getFullInfo(ycsxno);
                                    if (file != "")
                                    {
                                        string drawfilename = dt.Rows[0]["G_NAME"].ToString().Substring(0, 11) + "_" + dt.Rows[0]["G_CODE"].ToString().Substring(7, 1) + ".pdf";
                                        string pdffile = Dir + "\\BANVE\\" + drawfilename;
                                        ExcelFactory.editFileExcel(file, dt, cb1, saveycsxpath);
                                        if (inhaykhong == true)
                                        {
                                            if (File.Exists(pdffile))
                                            {
                                                new Form1().printPDF(pdffile);
                                            }
                                            else
                                            {
                                                MessageBox.Show("Không có bản vẽ : " + dt.Rows[0]["G_NAME"].ToString());
                                            }

                                        }
                                        startprogress = startprogress + 1;
                                        backgroundWorker1.ReportProgress(startprogress);
                                        //progressBar1.Value = startprogress;
                                    }

                                }
                                else
                                {
                                    dataGridView1.Rows[row.Index].Cells["PROD_REQUEST_NO"].Style.BackColor = Color.Red;
                                    MessageBox.Show("YCSX chưa được phê duyệt sử dụng");
                                }
                                

                            }

                        }

                    }
                    MessageBox.Show("Export Yêu cầu hoàn thành !");
                    progressBar1.Value = 0;
                    // MessageBox.Show(saveycsxpath);
                }
                



            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.ToString());
            }
        }


        public void print_PDF2(string printer_name, string pdf_path, string file_name)
        {
            foreach (string printerName in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
            {
               richTextBox1.Text +=printerName + "\n";
            }           
            IPrinter printer = new Printer();
            printer.PrintRawFile(printer_name, pdf_path, file_name);
        }
        public void print_PDF3()
        {
            foreach (string printerName in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
            {
                richTextBox1.Text += printerName + "\n";
            }

        }
        private void xuấtFileInYCSXToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    my_ycsx_flag = 0;
                    all_ycsx_flag = 0;
                    xuatfile_print_ycsx_flag = 1;
                    xuatfile_only_flag = 0;
                    xuatchithi_flag = 0;
                    checkbanve_flag = 0;
                    set_hoanthanh_flag = 0;
                    set_pending_flag = 0;
                    xoa_ycsx_flag = 0;
                    pheduyet_ycsx_flag = 0;
                    saveFolderPath = fbd.SelectedPath;
                    //MessageBox.Show(saveFolderPath);
                    if (!backgroundWorker1.IsBusy)
                    {
                        pictureBox1.Show();
                        backgroundWorker1.RunWorkerAsync();
                    }
                    else
                    {
                        MessageBox.Show("Tiến trình khác đang chạy, thử lại sau");
                    }

                }
            }
            //xuatfile_inycsx(true);
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if(checkBox4.Checked == true)
            {
                for (int r = 0; r < dataGridView1.Rows.Count; r++)
                {
                    dataGridView1.Rows[r].Cells["SELECT"].Value = true;
                }
            }
            else
            {
                for (int r = 0; r < dataGridView1.Rows.Count; r++)
                {
                    dataGridView1.Rows[r].Cells["SELECT"].Value = false;
                }
            }            
        }

        private void xuấtFileYCSXOnlyToolStripMenuItem_Click(object sender, EventArgs e)
        {
           

            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    my_ycsx_flag = 0;
                    all_ycsx_flag = 0;
                    xuatfile_print_ycsx_flag = 0;
                    xuatfile_only_flag = 1;
                    xuatchithi_flag = 0;
                    checkbanve_flag = 0;
                    set_hoanthanh_flag = 0;
                    set_pending_flag = 0;
                    xoa_ycsx_flag = 0;
                    pheduyet_ycsx_flag = 0;
                    saveFolderPath = fbd.SelectedPath;
                    //MessageBox.Show(saveFolderPath);
                    if (!backgroundWorker1.IsBusy)
                    {
                        pictureBox1.Show();
                        backgroundWorker1.RunWorkerAsync();
                    }
                    else
                    {
                        MessageBox.Show("Tiến trình khác đang chạy, thử lại sau");
                    }

                }
            }

            

            //xuatfile_inycsx(false);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            my_ycsx_flag = 1;
            all_ycsx_flag = 0;
            xuatfile_print_ycsx_flag = 0;
            xuatfile_only_flag = 0;
            xuatchithi_flag = 0;
            checkbanve_flag = 0;
            set_hoanthanh_flag = 0;
            set_pending_flag = 0;
            xoa_ycsx_flag = 0;
            pheduyet_ycsx_flag = 0;

            if (!backgroundWorker1.IsBusy)
            {
                pictureBox1.Show();
                backgroundWorker1.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Tiến trình khác đang chạy, thử lại sau");
            }

            /*
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();         
            dt = pro.getEmployeeName(Login_ID);
            string my_name = dt.Rows[0]["EMPL_NAME"].ToString();
            textBox5.Text = my_name;
            traYCSXManager();
            textBox5.Text = "";
            */
        }

        public void checkbanve()
        {
            dataGridView1.EndEdit();
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (!row.IsNewRow)
                {
                    if (row.Cells["SELECT"].Value.ToString() == "True")
                    {
                        string Dir = System.IO.Directory.GetCurrentDirectory();
                        string g_name_ = row.Cells["G_NAME"].Value.ToString();
                        string g_code_ = row.Cells["G_CODE"].Value.ToString();
                        string gname = g_name_.Substring(0, 11);
                        string gcode = g_code_.Substring(7, 1);

                        string drawpath = Dir + "\\BANVE\\" + gname + "_" + gcode + ".pdf";
                        //MessageBox.Show(drawpath);
                        if (!File.Exists(drawpath))
                        {
                            dataGridView1.Rows[row.Index].Cells["PROD_REQUEST_NO"].Style.BackColor = Color.Red;
                        }
                        else
                        {
                            dataGridView1.Rows[row.Index].Cells["PROD_REQUEST_NO"].Style.BackColor = Color.LightGreen;
                        }
                    }
                }
            }
            MessageBox.Show("Đã check bản vẽ xong, kiểm tra lại các dòng bôi đỏ");
        }
        private void checkBảnVẽToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            my_ycsx_flag = 0;
            all_ycsx_flag = 0;
            xuatfile_print_ycsx_flag = 0;
            xuatfile_only_flag = 0;
            xuatchithi_flag = 0;
            checkbanve_flag = 1;
            set_hoanthanh_flag = 0;
            set_pending_flag = 0;
            xoa_ycsx_flag = 0;
            pheduyet_ycsx_flag = 0;

            if (!backgroundWorker1.IsBusy)
            {
                pictureBox1.Show();
                backgroundWorker1.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Tiến trình khác đang chạy, thử lại sau");
            }

            /*

            dataGridView1.EndEdit();
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (!row.IsNewRow)
                {
                    if (row.Cells["SELECT"].Value.ToString() == "True")
                    {
                        string Dir = System.IO.Directory.GetCurrentDirectory();
                        string g_name_ = row.Cells["G_NAME"].Value.ToString();
                        string g_code_ = row.Cells["G_CODE"].Value.ToString();
                        string gname = g_name_.Substring(0, 11);
                        string gcode = g_code_.Substring(7, 1);

                        string drawpath = Dir + "\\BANVE\\" + gname + "_" + gcode + ".pdf";
                        //MessageBox.Show(drawpath);
                        if (!File.Exists(drawpath))
                        {
                            dataGridView1.Rows[row.Index].Cells["PROD_REQUEST_NO"].Style.BackColor = Color.Red;
                        }
                        else
                        {
                            dataGridView1.Rows[row.Index].Cells["PROD_REQUEST_NO"].Style.BackColor = Color.LightGreen;
                        }
                    }
                }
            }
            MessageBox.Show("Đã check bản vẽ xong, kiểm tra lại các dòng bôi đỏ");           */
        }

        public void xuatchithi()
        {
            CheckBox ckb = new CheckBox();
            ckb.Checked = false;
            dataGridView1.EndEdit();
            try
            {
                string Dir = System.IO.Directory.GetCurrentDirectory();
                //MessageBox.Show(Dir);
                string file = Dir + "\\templateqlsx.xlsx";
                string saveycsxpath = "";


                using (var fbd = new FolderBrowserDialog())
                {
                    DialogResult result = fbd.ShowDialog();

                    if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                    {
                        saveycsxpath = fbd.SelectedPath;

                        ProductBLL pro = new ProductBLL();
                        DataTable dt = new DataTable();

                        int checkRowsCount = 0;
                        for (int r = 0; r < dataGridView1.Rows.Count; r++)
                        {
                            DataGridViewRow row = dataGridView1.Rows[r];
                            if (!row.IsNewRow)
                            {
                                if ((Boolean)((DataGridViewCheckBoxCell)row.Cells["SELECT"]).FormattedValue)
                                {
                                    checkRowsCount++;
                                }

                            }
                        }
                        // MessageBox.Show("Số dòng đã chọn: " + checkRowsCount);
                        progressBar1.Minimum = 0; //Đặt giá trị nhỏ nhất cho ProgressBar
                        progressBar1.Maximum = checkRowsCount; //Đặt giá trị lớn nhất cho ProgressBar


                        int startprogress = 0;

                        for (int r = 0; r < dataGridView1.Rows.Count; r++)
                        {
                            DataGridViewRow row = dataGridView1.Rows[r];
                            if (!row.IsNewRow)
                            {
                                if ((Boolean)((DataGridViewCheckBoxCell)row.Cells["SELECT"]).FormattedValue)
                                {
                                    string ycsxno = row.Cells["PROD_REQUEST_NO"].Value.ToString();
                                    dt = pro.getFullInfo(ycsxno);
                                    if (file != "")
                                    {
                                        ExcelFactory.editFileExcelQLSX(file, dt, ckb, saveycsxpath);
                                        startprogress = startprogress + 1;
                                        backgroundWorker1.ReportProgress(startprogress);
                                        //progressBar1.Value = startprogress;
                                    }
                                }
                            }
                        }

                        MessageBox.Show("Export Yêu cầu và thêm chỉ thị hoàn thành !");
                        progressBar1.Value = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.ToString());
            }
        }

        public void xuatchithi2()
        {
            CheckBox ckb = new CheckBox();
            ckb.Checked = false;
            dataGridView1.EndEdit();
            try
            {
                string Dir = System.IO.Directory.GetCurrentDirectory();
                //MessageBox.Show(Dir);
                string file = Dir + "\\templateqlsx.xlsx";
                string saveycsxpath = "";              

                if (saveFolderPath != "")
                {
                    saveycsxpath = saveFolderPath;

                    ProductBLL pro = new ProductBLL();
                    DataTable dt = new DataTable();

                    int checkRowsCount = 0;
                    for (int r = 0; r < dataGridView1.Rows.Count; r++)
                    {
                        DataGridViewRow row = dataGridView1.Rows[r];
                        if (!row.IsNewRow)
                        {
                            if ((Boolean)((DataGridViewCheckBoxCell)row.Cells["SELECT"]).FormattedValue)
                            {
                                checkRowsCount++;
                            }

                        }
                    }
                    // MessageBox.Show("Số dòng đã chọn: " + checkRowsCount);
                    progressBar1.Minimum = 0; //Đặt giá trị nhỏ nhất cho ProgressBar
                    progressBar1.Maximum = checkRowsCount; //Đặt giá trị lớn nhất cho ProgressBar


                    int startprogress = 0;

                    for (int r = 0; r < dataGridView1.Rows.Count; r++)
                    {
                        DataGridViewRow row = dataGridView1.Rows[r];
                        if (!row.IsNewRow)
                        {
                            string PDUYET = row.Cells["PDUYET"].Value.ToString();
                            if (PDUYET == "1")
                            {
                                if ((Boolean)((DataGridViewCheckBoxCell)row.Cells["SELECT"]).FormattedValue)
                                {
                                    string ycsxno = row.Cells["PROD_REQUEST_NO"].Value.ToString();
                                    dt = pro.getFullInfo(ycsxno);
                                    if (file != "")
                                    {
                                        ExcelFactory.editFileExcelQLSX(file, dt, ckb, saveycsxpath);
                                        startprogress = startprogress + 1;
                                        backgroundWorker1.ReportProgress(startprogress);
                                        //progressBar1.Value = startprogress;
                                    }
                                }
                            }
                            else
                            {
                                dataGridView1.Rows[row.Index].Cells["PROD_REQUEST_NO"].Style.BackColor = Color.Red;
                                MessageBox.Show("YCSX chưa được phê duyệt nên không xuất chỉ thị cho yc này");
                            }
                        }
                    }

                    MessageBox.Show("Export Yêu cầu và thêm chỉ thị hoàn thành !");
                    progressBar1.Value = 0;
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.ToString());
            }
        }

        private void xuấtChỉThịToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    my_ycsx_flag = 0;
                    all_ycsx_flag = 0;
                    xuatfile_print_ycsx_flag = 0;
                    xuatfile_only_flag = 0;
                    xuatchithi_flag = 1;
                    checkbanve_flag = 0;
                    set_hoanthanh_flag = 0;
                    set_pending_flag = 0;
                    xoa_ycsx_flag = 0;
                    pheduyet_ycsx_flag = 0;

                    saveFolderPath = fbd.SelectedPath;
                    //MessageBox.Show(saveFolderPath);
                    if (!backgroundWorker1.IsBusy)
                    {
                        pictureBox1.Show();
                        backgroundWorker1.RunWorkerAsync();
                    }
                    else
                    {
                        MessageBox.Show("Tiến trình khác đang chạy, thử lại sau");
                    }

                }
            }

            /*
            CheckBox ckb = new CheckBox();
            ckb.Checked = false;
            dataGridView1.EndEdit();
            try
            {
                string Dir = System.IO.Directory.GetCurrentDirectory();
                //MessageBox.Show(Dir);
                string file = Dir + "\\templateqlsx.xlsx";
                string saveycsxpath = "";


                using (var fbd = new FolderBrowserDialog())
                {
                    DialogResult result = fbd.ShowDialog();

                    if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                    {
                        saveycsxpath = fbd.SelectedPath;

                        ProductBLL pro = new ProductBLL();
                        DataTable dt = new DataTable();

                        int checkRowsCount = 0;
                        for (int r = 0; r < dataGridView1.Rows.Count; r++)
                        {
                            DataGridViewRow row = dataGridView1.Rows[r];                           
                            if (!row.IsNewRow)
                            {                                
                                if ((Boolean)((DataGridViewCheckBoxCell)row.Cells["SELECT"]).FormattedValue)
                                {
                                    checkRowsCount++;
                                }

                            }
                        }                       
                       // MessageBox.Show("Số dòng đã chọn: " + checkRowsCount);
                        progressBar1.Minimum = 0; //Đặt giá trị nhỏ nhất cho ProgressBar
                        progressBar1.Maximum = checkRowsCount; //Đặt giá trị lớn nhất cho ProgressBar


                        int startprogress = 0;

                        for (int r = 0; r < dataGridView1.Rows.Count; r++)
                        {
                            DataGridViewRow row = dataGridView1.Rows[r];
                            if (!row.IsNewRow)
                            {                                
                                if ((Boolean)((DataGridViewCheckBoxCell)row.Cells["SELECT"]).FormattedValue)
                                {
                                    string ycsxno = row.Cells["PROD_REQUEST_NO"].Value.ToString();
                                    dt = pro.getFullInfo(ycsxno);
                                    if (file != "")
                                    {
                                        ExcelFactory.editFileExcelQLSX(file, dt, ckb, saveycsxpath);
                                        startprogress = startprogress + 1;
                                        progressBar1.Value = startprogress;
                                    }
                                }
                            }
                        }
                        
                        MessageBox.Show("Export Yêu cầu và thêm chỉ thị hoàn thành !");
                        progressBar1.Value = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.ToString());
            }

            */
        }


        public void xoa_ycsx()
        {
            dataGridView1.EndEdit();
            if (textBox3.Text == "xoa")
            {
                if (MessageBox.Show("Bạn thực sự muốn xóa YCSX?", "Xóa YCSX ?", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    try
                    {
                        ProductBLL pro = new ProductBLL();
                        DataTable dt = new DataTable();


                        int checkRowsCount = 0;
                        for (int r = 0; r < dataGridView1.Rows.Count; r++)
                        {
                            DataGridViewRow row = dataGridView1.Rows[r];
                            if (!row.IsNewRow)
                            {
                                if ((Boolean)((DataGridViewCheckBoxCell)row.Cells["SELECT"]).FormattedValue)
                                {
                                    checkRowsCount++;
                                }
                            }
                        }
                        progressBar1.Minimum = 0; //Đặt giá trị nhỏ nhất cho ProgressBar
                        progressBar1.Maximum = checkRowsCount; //Đặt giá trị lớn nhất cho ProgressBar
                        int startprogress = 0;

                        for (int r = 0; r < dataGridView1.Rows.Count; r++)
                        {
                            DataGridViewRow row = dataGridView1.Rows[r];
                            if (!row.IsNewRow)
                            {
                                if ((Boolean)((DataGridViewCheckBoxCell)row.Cells["SELECT"]).FormattedValue)
                                {
                                    string ycsxno = row.Cells["PROD_REQUEST_NO"].Value.ToString();
                                    dt = pro.checkYCSXO300(ycsxno);
                                    if (dt.Rows.Count == 0)
                                    {
                                        dt = pro.DeleteYCSX(ycsxno);
                                        pro.writeHistory("002", Login_ID, "YCSX TABLE", "XOA", "XOA YCSX", "0");
                                    }
                                    startprogress = startprogress + 1;
                                    backgroundWorker1.ReportProgress(startprogress);
                                    //progressBar1.Value = startprogress;
                                }
                            }
                        }
                        //progressBar1.Value = 0;
                        dataGridView1.ClearSelection();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }
            }
            else
            {
                MessageBox.Show("Lêu lêu, còn lâu mới xóa được !");
            }
        }

        public void pheduyet_ycsx()
        {
            dataGridView1.EndEdit();
            if (textBox3.Text == "pd" || Login_ID=="LVT1906")
            {
                if (MessageBox.Show("Bạn thực sự muốn phê duyệt YCSX?", "Phê duyệt YCSX ?", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    try
                    {
                        ProductBLL pro = new ProductBLL();
                        DataTable dt = new DataTable();


                        int checkRowsCount = 0;
                        for (int r = 0; r < dataGridView1.Rows.Count; r++)
                        {
                            DataGridViewRow row = dataGridView1.Rows[r];
                            if (!row.IsNewRow)
                            {
                                if ((Boolean)((DataGridViewCheckBoxCell)row.Cells["SELECT"]).FormattedValue)
                                {
                                    checkRowsCount++;
                                }
                            }
                        }
                        progressBar1.Minimum = 0; //Đặt giá trị nhỏ nhất cho ProgressBar
                        progressBar1.Maximum = checkRowsCount; //Đặt giá trị lớn nhất cho ProgressBar
                        int startprogress = 0;

                        for (int r = 0; r < dataGridView1.Rows.Count; r++)
                        {
                            DataGridViewRow row = dataGridView1.Rows[r];
                            if (!row.IsNewRow)
                            {
                                if ((Boolean)((DataGridViewCheckBoxCell)row.Cells["SELECT"]).FormattedValue)
                                {
                                    string ycsxno = row.Cells["PROD_REQUEST_NO"].Value.ToString();
                                        dt = pro.PheDuyetYCSX(ycsxno); 
                                    startprogress = startprogress + 1;
                                    backgroundWorker1.ReportProgress(startprogress);
                                    //progressBar1.Value = startprogress;
                                }
                            }
                        }
                        //progressBar1.Value = 0;
                        dataGridView1.ClearSelection();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }
            }
            else
            {
                MessageBox.Show("Không đủ quyền hạn để phê duyệt !");
            }
        }


        private void xóaYCSXToolStripMenuItem_Click(object sender, EventArgs e)
        {
            my_ycsx_flag = 0;
            all_ycsx_flag = 0;
            xuatfile_print_ycsx_flag = 0;
            xuatfile_only_flag = 0;
            xuatchithi_flag = 0;
            checkbanve_flag = 0;
            set_hoanthanh_flag = 0;
            set_pending_flag = 0;
            xoa_ycsx_flag = 1;
            pheduyet_ycsx_flag = 0;

            if (!backgroundWorker1.IsBusy)
            {
                pictureBox1.Show();
                backgroundWorker1.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Tiến trình khác đang chạy, thử lại sau");
            }

            /*
            dataGridView1.EndEdit();
            if (textBox3.Text == "xoa")
            {
                if (MessageBox.Show("Bạn thực sự muốn xóa YCSX?", "Xóa YCSX ?", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    try
                    {
                        ProductBLL pro = new ProductBLL();
                        DataTable dt = new DataTable();

                     
                        int checkRowsCount = 0;
                        for (int r = 0; r < dataGridView1.Rows.Count; r++)
                        {
                            DataGridViewRow row = dataGridView1.Rows[r];
                            if (!row.IsNewRow)
                            {
                                if ((Boolean)((DataGridViewCheckBoxCell)row.Cells["SELECT"]).FormattedValue)
                                {
                                    checkRowsCount++;
                                }

                            }
                        }

                        progressBar1.Minimum = 0; //Đặt giá trị nhỏ nhất cho ProgressBar
                        progressBar1.Maximum = checkRowsCount; //Đặt giá trị lớn nhất cho ProgressBar
                        int startprogress = 0;

                        for (int r = 0; r < dataGridView1.Rows.Count; r++)
                        {
                            DataGridViewRow row = dataGridView1.Rows[r];
                            if (!row.IsNewRow)
                            {
                                if ((Boolean)((DataGridViewCheckBoxCell)row.Cells["SELECT"]).FormattedValue)
                                {
                                    string ycsxno = row.Cells["PROD_REQUEST_NO"].Value.ToString();
                                    dt = pro.checkYCSXO300(ycsxno);
                                    if (dt.Rows.Count == 0)
                                    {
                                        dt = pro.DeleteYCSX(ycsxno);
                                        pro.writeHistory("002", Login_ID, "YCSX TABLE", "XOA", "XOA YCSX", "0");
                                    }
                                    startprogress = startprogress + 1;
                                    progressBar1.Value = startprogress;
                                }                                
                            }
                        }                          
                        
                        progressBar1.Value = 0;
                        dataGridView1.ClearSelection();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }
            }
            else
            {
                MessageBox.Show("Lêu lêu, còn lâu mới xóa được !");
            }

            */
        }

        private void testInPdfToolStripMenuItem_Click(object sender, EventArgs e)
        {
            print_PDF3();
        }

        private void uploadDataAmazoneChoYCSXToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AmazoneForm amazoneForm = new AmazoneForm();
            amazoneForm.Login_ID = Login_ID;
            dataGridView1.EndEdit();
            for (int r = 0; r < dataGridView1.Rows.Count; r++)
            {
                DataGridViewRow row = dataGridView1.Rows[r];
                if (!row.IsNewRow)
                {
                    if ((Boolean)((DataGridViewCheckBoxCell)row.Cells["SELECT"]).FormattedValue)
                    {
                        amazoneForm.PROD_REQUEST_NO = row.Cells["PROD_REQUEST_NO"].Value.ToString();
                    }
                }
            }
            amazoneForm.Show();
        }

        private void updataAmazoneToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AmazoneForm amazoneForm = new AmazoneForm();
            amazoneForm.Login_ID = Login_ID;
            dataGridView1.EndEdit();
            for (int r = 0; r < dataGridView1.Rows.Count; r++)
            {
                DataGridViewRow row = dataGridView1.Rows[r];
                if (!row.IsNewRow)
                {
                    if ((Boolean)((DataGridViewCheckBoxCell)row.Cells["SELECT"]).FormattedValue)
                    {
                        amazoneForm.PROD_REQUEST_NO = row.Cells["PROD_REQUEST_NO"].Value.ToString();
                    }
                }
            }
            amazoneForm.Show();
        }


        public void my_ycsx_Async()
        {
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dt = pro.getEmployeeName(Login_ID);
            string my_name = dt.Rows[0]["EMPL_NAME"].ToString();
            textBox5.Text = my_name;
            traYCSXManager();
            textBox5.Text = "";
        }
        public void all_ycsx_Async()
        {  
            traYCSXManager();
        }

        public void xuatfile_print_ycsx_Async()
        {
            xuatfile_inycsx2(true);
        }

        public void xuatfile_only_Async()
        {
            xuatfile_inycsx2(false);
        }

        public void xuatchithi_Async()
        {
            xuatchithi2();
        }

        public void checkbanve_Async()
        {
            checkbanve();
        }
        
        public void sethoanthanh_Async()
        {
            sethoanthanh();
        }

        public void setpending_Async()
        {
            setpending();
        }

        public void xoa_ycsx_Async()
        {
            xoa_ycsx();
        }
        public void pheduyet_ycsx_Async()
        {
            pheduyet_ycsx();
        }
        public void xuly_ycsx_mamager()
        {
            if(my_ycsx_flag == 1)
            {
                my_ycsx_Async();

            }
            else if(all_ycsx_flag == 1)
            {
                all_ycsx_Async();
            }
            else if (xuatfile_print_ycsx_flag == 1)
            {
                xuatfile_print_ycsx_Async();
            }
            else if (xuatfile_only_flag == 1)
            {
                xuatfile_only_Async();
            }
            else if (xuatchithi_flag == 1)
            {
                xuatchithi_Async();
            }
            else if (checkbanve_flag == 1)
            {
                checkbanve_Async();
            }
            else if (set_hoanthanh_flag == 1)
            {
                sethoanthanh_Async();
            }
            else if (set_pending_flag == 1)
            {
                setpending_Async();
            }
            else if (xoa_ycsx_flag == 1)
            {
                xoa_ycsx_Async();
            }
            else if (pheduyet_ycsx_flag == 1)
            {
                pheduyet_ycsx_Async();
            }



        }
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            xuly_ycsx_mamager();
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (my_ycsx_flag == 1)
            {

            }
            else if (all_ycsx_flag == 1)
            {

            }
            else if (xuatfile_print_ycsx_flag == 1)
            {
                progressBar1.Value = e.ProgressPercentage;
            }
            else if (xuatfile_only_flag == 1)
            {
                progressBar1.Value = e.ProgressPercentage;
            }
            else if (xuatchithi_flag == 1)
            {
                progressBar1.Value = e.ProgressPercentage;
            }
            else if (checkbanve_flag == 1)
            {

            }
            else if (set_hoanthanh_flag == 1)
            {
                progressBar1.Value = e.ProgressPercentage;
            }
            else if (set_pending_flag == 1)
            {
                progressBar1.Value = e.ProgressPercentage;
            }
            else if (xoa_ycsx_flag == 1)
            {
                progressBar1.Value = e.ProgressPercentage;
            }

        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            pictureBox1.Hide();

            if (my_ycsx_flag == 1)
            {
                dataGridView1.DataSource = dtgv1_data;
                setRowNumber(dataGridView1);
                formatYCSXTable(dataGridView1);
                changeHeaderText(dataGridView1);

                DataGridViewCheckBoxColumn ck = new DataGridViewCheckBoxColumn();
                ck.Name = "SELECT";
                ck.HeaderText = "CHỌN";
                ck.Width = 50;
                ck.ReadOnly = false;
                dataGridView1.Columns.Insert(0, ck);
                for (int r = 0; r < dataGridView1.Rows.Count; r++)
                {
                    dataGridView1.Rows[r].Cells["SELECT"].Value = false;
                }
                dataGridView1.Columns["PDUYET"].ReadOnly = true;
                MessageBox.Show("Đã loát: " + dtgv1_data.Rows.Count + " dòng");

            }
            else if (all_ycsx_flag == 1)
            {
                dataGridView1.DataSource = dtgv1_data;
                setRowNumber(dataGridView1);
                formatYCSXTable(dataGridView1);
                changeHeaderText(dataGridView1);

                DataGridViewCheckBoxColumn ck = new DataGridViewCheckBoxColumn();
                ck.Name = "SELECT";
                ck.HeaderText = "CHỌN";
                ck.Width = 50;
                ck.ReadOnly = false;
                dataGridView1.Columns.Insert(0, ck);
                for (int r = 0; r < dataGridView1.Rows.Count; r++)
                {
                    dataGridView1.Rows[r].Cells["SELECT"].Value = false;
                }
                dataGridView1.Columns["PDUYET"].ReadOnly = true;
                MessageBox.Show("Đã loát: " + dtgv1_data.Rows.Count + " dòng");
            }
            else if (xuatfile_print_ycsx_flag == 1)
            {

            }
            else if (xuatfile_only_flag == 1)
            {

            }
            else if (xuatchithi_flag == 1)
            {

            }
            else if (checkbanve_flag == 1)
            {

            }
            else if (set_hoanthanh_flag == 1)
            {

            }
            else if (set_pending_flag == 1)
            {

            }
            else if (xoa_ycsx_flag == 1)
            {

            }

        }

        private void phêtDuyệtYCSXToolStripMenuItem_Click(object sender, EventArgs e)
        {
            my_ycsx_flag = 0;
            all_ycsx_flag = 0;
            xuatfile_print_ycsx_flag = 0;
            xuatfile_only_flag = 0;
            xuatchithi_flag = 0;
            checkbanve_flag = 0;
            set_hoanthanh_flag = 0;
            set_pending_flag = 0;
            xoa_ycsx_flag = 0;
            pheduyet_ycsx_flag = 1;

            if (!backgroundWorker1.IsBusy)
            {
                pictureBox1.Show();
                backgroundWorker1.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Tiến trình khác đang chạy, thử lại sau");
            }
        }
    }
}
