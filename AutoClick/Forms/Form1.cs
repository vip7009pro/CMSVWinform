using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Security.Cryptography;
using System.Data.SqlClient;
using System.IO;
using eX = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using Microsoft.SqlServer.Server;
using System.Globalization;
using System.Reflection;
using AutoUpdaterDotNET;
using Microsoft.VisualBasic;

namespace AutoClick
{
    
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            timer1.Start();
            pictureBox1.Hide();
            checkUpdate();
            int h = Screen.PrimaryScreen.WorkingArea.Height;
            int w = Screen.PrimaryScreen.WorkingArea.Width;
            this.ClientSize = new Size(w, h);

            this.dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
          
            this.ContextMenuStrip = contextMenuStrip1;
            dataGridView1.MultiSelect = true;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            if (!System.Windows.Forms.SystemInformation.TerminalServerSession)
            {
                Type dgvType = dataGridView1.GetType();
                PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                  BindingFlags.Instance | BindingFlags.NonPublic);
                pi.SetValue(dataGridView1, true, null);
            }
            
            //dataGridView1.Hide();            
            dateTimePicker1.Value = DateTime.Today.AddDays(0);
            dateTimePicker2.Value = DateTime.Today.AddDays(0);
            button4.Enabled = false;
            //button5.Enabled = false;
            button3.Enabled = false;
            button8.Enabled = false;
            //button1.Enabled = false;
            button16.Enabled = false;

            button14.Enabled = false;
            //MessageBox.Show("Xin chào " + LoginID);
            label6.Text = "Xin chào " + LoginID;
            //checkBox2.Checked = true;             
            //checkKinhDoanhvsKiemTraG_CODE();
        }

        public void checkKinhDoanhvsKiemTraG_CODE()
        {
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dt = pro.testQuery("SELECT M100.G_NAME, KIEMTRA.G_CODE AS KIEMTRA_CODE, isnull(KINHDOANH.G_CODE,'?') AS KINHDOANH_CODE FROM (SELECT DISTINCT G_CODE FROM ZTBINSPECTNGTB) AS KIEMTRA LEFT JOIN (SELECT DISTINCT G_CODE FROM ZTBPOTable) AS KINHDOANH ON (KIEMTRA.G_CODE = KINHDOANH.G_CODE) LEFT JOIN M100 ON (M100.G_CODE = KIEMTRA.G_CODE) WHERE KINHDOANH.G_CODE is null");
            dataGridView1.DataSource = dt;
            MessageBox.Show("Có " + dt.Rows.Count.ToString() + " code Có vào phòng kiểm tra mà không có PO");
        }
        public void checkUpdate()
        {
            System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
            System.Diagnostics.FileVersionInfo fvi = System.Diagnostics.FileVersionInfo.GetVersionInfo(assembly.Location);
            AutoUpdater.CheckForUpdateEvent += AutoUpdaterOnCheckForUpdateEvent;
            string version = fvi.FileVersion;
            label21.Text = "Phiên bản: " + version;
            AutoUpdater.DownloadPath = "update";
            AutoUpdater.Start("http://14.160.33.94:3010/update/update.xml");

            System.Timers.Timer timer = new System.Timers.Timer
            {
                Interval = 1 * 60 * 1000,
                SynchronizingObject = this
            };
            timer.Elapsed += delegate
            {
                AutoUpdater.Start("http://14.160.33.94:3010/update/update.xml");
            };
            timer.Start();
        }

        private void AutoUpdaterOnCheckForUpdateEvent(UpdateInfoEventArgs args)
        {
            if (args.IsUpdateAvailable)
            {
                DialogResult dialogResult;
                dialogResult =
                        MessageBox.Show(
                            $@"Bạn ơi, phần mềm của bạn có phiên bản mới {args.CurrentVersion}. Phiên bản bạn đang sử dụng hiện tại  {args.InstalledVersion}. Bạn có muốn cập nhật phần mềm không?", @"Cập nhật phần mềm",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Information);

                if (dialogResult.Equals(DialogResult.Yes) || dialogResult.Equals(DialogResult.OK))
                {
                    try
                    {
                        if (AutoUpdater.DownloadUpdate(args))
                        {
                            Application.Exit();
                        }
                    }
                    catch (Exception exception)
                    {
                        MessageBox.Show(exception.Message, exception.GetType().ToString(), MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                    }
                }
            }
            else
            {
               /* MessageBox.Show(@"Phiên bản bạn đang sử dụng đã được cập nhật mới nhất.", @"Cập nhật phần mềm",
                    MessageBoxButtons.OK, MessageBoxIcon.Information); */
            }
        }

        public int po_flag = 0, invoice_flag = 0, ycsx_flag = 0, fcst_flag = 0, khgh_flag = 0, bom_flag=0; 
        public int import_excel_flag = 0;
        public int check_po_flag = 0;
        public int up_po_flag = 0;
        public int check_invoice_flag = 0;
        public int up_invoice_flag = 0;
        public int check_fcst_flag = 0;
        public int up_fcst_flag = 0;
        public int check_plan_flag = 0;
        public int up_plan_flag = 0;

        public string LoginID = "NBT1901";
        public List<string> listchuabanve = null;
        List<YeuCauSanXuat> dsNV = null;
        public string[] monthArray = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L" };
        public string[] dayArray = { "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V" };
 
        public int checkDate(DateTime dateTime)
        {
            if (dateTime <= DateTime.Today)
            {
                return 1;
            }
            else
            {
                return 0;
            }            
        }
        public int checkInvoicevsPODate(DateTime inVoicedateTime, DateTime PODateTime)
        {
            if (inVoicedateTime >= PODateTime) 
            {
                return 1;                
            }
            else
            {                
                return 0;               
            }

        }

        public void setRowNumber(DataGridView dgv)
        {
            foreach (DataGridViewRow row in dgv.Rows)
            {
                row.HeaderCell.Value = (row.Index + 1).ToString();
            }
        }
        public void CopyToClipboardWithHeaders(DataGridView _dgv)
        {   //Copy to clipboard
            _dgv.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            DataObject dataObj = _dgv.GetClipboardContent();
            if (dataObj != null)
            Clipboard.SetDataObject(dataObj);
        }

        public string returnUser()
        {
            return LoginID;
        }

        public void printPDF(string pdffilepath)
        {            
            ProcessStartInfo info = new ProcessStartInfo();
            info.Verb = "print";
            info.FileName = pdffilepath;
            info.CreateNoWindow = true;
            info.WindowStyle = ProcessWindowStyle.Normal;
            Process p = new Process();
            p.StartInfo = info;
            p.Start();           
            if (!p.WaitForExit(500))
            {
                p.Kill();
            }
            
        }

        public string STYMD(int y, int m, int d)
        {
            string ymd,sty,stm,std;
            sty = "" + y;
            stm = "" + m;
            std = "" + d;
            if (m<10)
            {
                stm = "0" + m;
            }
            if(d<10)
            {
                std = "0" + d;
            }
            ymd = sty + stm + std;
            return ymd;
        }


        public string STYMD2(int y, int m, int d)
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
            ymd = sty + "-" + stm + "-" + std;
            return ymd;
        }


        private void button1_Click(object sender, EventArgs e)
        {
            string query = "WHERE ";
            string fromdate, todate;
            fromdate = STYMD(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day);
            todate = STYMD(dateTimePicker2.Value.Year, dateTimePicker2.Value.Month, dateTimePicker2.Value.Day);
            string ngaythang = "PO_DATE BETWEEN '" + fromdate + "' AND '" + todate + "' ";

            string code;            
            if(textBox1.Text != "")
            {
                code = "AND G_NAME LIKE '%"+ textBox1.Text + "%' ";
            }
            else
            {
                code = "";
            }

            string empl_name = "";
            if(textBox8.Text!="")
            {
                empl_name = "AND EMPL_NAME LIKE '%" + textBox8.Text + "%' ";
            }
            else
            {
                empl_name = "";
            }

            string cust_name_kd = "";
            if (textBox9.Text != "")
            {
                cust_name_kd = "AND CUST_NAME_KD LIKE '%" + textBox9.Text + "%' ";
            }
            else
            {
                cust_name_kd = "";
            }
            string prod_type = "";

            if (textBox10.Text != "")
            {
                prod_type = "AND PROD_TYPE LIKE '%" + textBox10.Text + "%' ";
            }
            else
            {
                prod_type = "";
            }

            string po_no = "";

            if (textBox11.Text != "")
            {
                po_no = "AND PO_NO='" + textBox11.Text + "' ";
            }
            else
            {
                po_no = "";
            }

            string material = "";

            if (textBox12.Text != "")
            {
                material = "AND PROD_MAIN_MATERIAL LIKE '%" + textBox12.Text + "%' ";
            }
            else
            {
                material = "";
            }

            string overdue = "";

            if (textBox13.Text != "")
            {
                overdue = "AND OVERDUE LIKE '%" + textBox13.Text + "%' ";
            }
            else
            {
                overdue = "";
            }

            query += ngaythang + code + empl_name + cust_name_kd + prod_type + po_no + material + overdue;
            MessageBox.Show(query);
        }


        public string CreateHeader2()
        {
            ProductBLL pro1 = new ProductBLL();
            String ngaygiohethong = pro1.getsystemDateTime();
            String ngayhethong = ngaygiohethong.Substring(0, 4) + ngaygiohethong.Substring(5, 2) + ngaygiohethong.Substring(8, 2);

            string year = ngaygiohethong.Substring(3, 1);
            int month = int.Parse(ngaygiohethong.Substring(5, 2));
            int day = int.Parse(ngaygiohethong.Substring(8, 2));
            string monthstr = monthArray[month - 1];
            string daystr = dayArray[day - 1];
            string header = year + monthstr + daystr;
            return header;
        }


        public string CreateHeader()
        {
            ProductBLL pro1 = new ProductBLL();
            String ngaygiohethong = pro1.getsystemDateTime();
            String ngayhethong = ngaygiohethong.Substring(0, 4) + ngaygiohethong.Substring(5, 2) + ngaygiohethong.Substring(8, 2);

            string year = DateTime.Today.Year.ToString().Substring(3);
            int month = int.Parse(DateTime.Today.Month.ToString());
            int day = int.Parse(DateTime.Today.Day.ToString());
            string monthstr = monthArray[month - 1];
            string daystr = dayArray[day - 1];
            string header = year + monthstr + daystr;
            return header;
        }

        static Random rd = new Random();
        internal static string CreateString(int stringLength)
        {
            const string allowedChars = "ABCDEFGHJKLMNOPQRSTUVWXYZ";
            char[] chars = new char[stringLength];

            for (int i = 0; i < stringLength; i++)
            {
                chars[i] = allowedChars[rd.Next(0, allowedChars.Length)];
            }
            return new string(chars);
        }

        internal static string CreateNumber(int stringLength)
        {
            const string allowedChars = "0123456789";
            char[] chars = new char[stringLength];

            for (int i = 0; i < stringLength; i++)
            {
                chars[i] = allowedChars[rd.Next(0, allowedChars.Length)];
            }
            return new string(chars);
        }

        internal static string GetRandomString(int stringLength)
        {
            StringBuilder sb = new StringBuilder();
            int numGuidsToConcat = (((stringLength - 1) / 32) + 1);
            for (int i = 1; i <= numGuidsToConcat; i++)
            {
                sb.Append(Guid.NewGuid().ToString("N"));
            }
            return sb.ToString(0, stringLength);
        }




        public void Search(string item)
        {
            try
            {
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();
                dt = pro.downLoadYCSX(textBox1.Text);
                if (dt.Rows.Count > 0)
                {
                    dataGridView1.DataSource = dt;
                }
                else
                {
                    MessageBox.Show("Loi");
                }

              
                string temp_ycsx;
                temp_ycsx = GetRandomString(7).ToUpper();
                while (pro.checkDuplicateYCSX(temp_ycsx).Rows.Count > 0)
                {
                    temp_ycsx = GetRandomString(7).ToUpper();
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi !\n" + ex.ToString());
            }
        }



        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.ShowDialog();
                string file = ofd.FileName;

                if (file != "")
                {
                    dsNV = ExcelFactory.readFromExcelFile(file);                    
                    dataGridView1.DataSource = dsNV;
                    button4.Enabled = true;
                    button3.Enabled = true;
                    button16.Enabled = true;
                    /*
                    dataGridView1.Columns[0].HeaderText = "Mã Nhân Viên";
                    dataGridView1.Columns[1].HeaderText = "Tên Nhân Viên";
                    dataGridView1.Columns[2].HeaderText = "Tuổi Nhân Viên";
                    */

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (dataGridView1 == null || dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("Không có dữ liệu!");
            }
            else
            {
                ExcelFactory.writeToExcelFile(dataGridView1);
            }
            
        }

        private void button4_Click(object sender, EventArgs e)
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
                            string PROD_REQUEST_DATE, CODE_50, CODE_55, G_CODE, RIV_NO, PROD_REQUEST_QTY, CUST_CD, EMPL_NO, REMK, DELIVERY_DT;
                            string CTR_CD = "002";
                            string CODE_03 = "01";
                            int TP = 0, BTP = 0, CK = 0, W1 = 0, W2 = 0, W3 = 0, W4 = 0, W5 = 0, W6 = 0, W7 = 0, W8 = 0, PO_BALANCE = 0, TOTAL_FCST = 0, PDuyet = 0;
                            PROD_REQUEST_DATE = Convert.ToString(row.Cells[0].Value);
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
                                    if (lastycsxno.Substring(0, 3) != CreateHeader())
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
                                    else if (yccuoiint<100)
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

                                    string PROD_REQUEST_NO = CreateHeader() + lastycsxno;
                                    //pro.InsertYCSX(CTR_CD, PROD_REQUEST_DATE, PROD_REQUEST_NO, CODE_50, CODE_03, CODE_55, G_CODE, RIV_NO, PROD_REQUEST_QTY, CUST_CD, EMPL_NO, REMK, EMPL_NO, EMPL_NO, DELIVERY_DT);


                                }
                                else
                                {
                                    MessageBox.Show("Loi");
                                }

                                

                            }


                            /*
                            for (int i = 0; i < dataGridView1.Columns.Count; i++)
                            {
                                String header = dataGridView1.Columns[i].HeaderText;
                                String cellText = Convert.ToString(row.Cells[i].Value);

                                //MessageBox.Show(cellText);
                            }
                            */


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

        private void button5_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Bạn thực sự muốn xóa YCSX?", "Xóa YCSX ?", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                try
                {
                    ProductBLL pro = new ProductBLL();
                    DataTable dt = new DataTable();

                    var selectedRows = dataGridView1.SelectedRows
                   .OfType<DataGridViewRow>()
                   .Where(row => !row.IsNewRow)
                   .ToArray();

                    progressBar1.Minimum = 0; //Đặt giá trị nhỏ nhất cho ProgressBar
                    progressBar1.Maximum = dataGridView1.SelectedRows.Count; //Đặt giá trị lớn nhất cho ProgressBar
                    int startprogress = 0;

                    foreach (var row in selectedRows)
                    {
                        string ycsxno = row.Cells[3].Value.ToString();
                        dt = pro.DeleteYCSX(ycsxno);
                        startprogress = startprogress + 1;
                        label4.Text = "Progress: " + startprogress + "/" + dataGridView1.SelectedRows.Count;
                        progressBar1.Value = startprogress;
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

        public void traYCSX()
        {
            try
            {
                bom_flag = 0;
                fcst_flag = 0;
                invoice_flag = 0;
                khgh_flag = 0;
                po_flag = 0;
                ycsx_flag = 1;
                button4.Enabled = true;
                string fromdate, todate;
                fromdate = STYMD(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day);
                todate = STYMD(dateTimePicker2.Value.Year, dateTimePicker2.Value.Month, dateTimePicker2.Value.Day);
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();
                dt = pro.traYCSX(textBox1.Text, fromdate, todate);
                if (dt.Rows.Count > 0)
                {
                    dataGridView1.DataSource = dt;
                    setRowNumber(dataGridView1);
                    button3.Enabled = true;
                    MessageBox.Show("Đã load " + dt.Rows.Count + " dòng");
                    //button5.Enabled = true;
                }
                else
                {
                    MessageBox.Show("Không có kết quả nào");
                }
                button8.Enabled = true;

            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi kết nối ! " + ex.ToString());
            }
        }
        private void button6_Click(object sender, EventArgs e)
        {
            SOPFULL sopfull = new SOPFULL();
            sopfull.Show(); 
        }
        

        private void button7_Click(object sender, EventArgs e)
        {
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dt = pro.checkOutMaterial(returnLotArray(dataGridView1));
            dataGridView1.DataSource = dt;
            MessageBox.Show("Đã load : " + dt.Rows.Count.ToString() + " dòng");
            button3.Enabled = true;
        }

        private void button8_Click(object sender, EventArgs e)
        {
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

                        var selectedRows = dataGridView1.SelectedRows
                       .OfType<DataGridViewRow>()
                       .Where(row => !row.IsNewRow)
                       .ToArray();

                        progressBar1.Minimum = 0; //Đặt giá trị nhỏ nhất cho ProgressBar
                        progressBar1.Maximum = dataGridView1.SelectedRows.Count; //Đặt giá trị lớn nhất cho ProgressBar
                        int startprogress = 0;                      


                        foreach (var row in selectedRows)
                        {
                            string ycsxno = row.Cells[3].Value.ToString();
                            dt = pro.getFullInfo(ycsxno);
                            if (file != "")
                            {
                                ExcelFactory.editFileExcel(file, dt, checkBox1, saveycsxpath);
                                startprogress = startprogress + 1;
                                label4.Text = "Progress: " +  startprogress + "/" + dataGridView1.SelectedRows.Count;
                                progressBar1.Value = startprogress;
                                
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
            
            /*
            

            */
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            //MessageBox.Show(checkBox1.Checked.ToString());
        }

        private void button9_Click(object sender, EventArgs e)
        {
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dt = pro.getProductInfo(textBox1.Text);
            dataGridView1.DataSource = dt;
            MessageBox.Show("Đã load : " + dt.Rows.Count.ToString() + " dòng");
            button3.Enabled = true;
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            try
            {
                button4.Enabled = true;
                string fromdate, todate;
                fromdate = STYMD(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day);
                todate = STYMD(dateTimePicker2.Value.Year, dateTimePicker2.Value.Month, dateTimePicker2.Value.Day);
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();                
                dt = pro.traYCSXPIC(LoginID, textBox1.Text, fromdate, todate);
                if (dt.Rows.Count > 0)
                {
                    dataGridView1.DataSource = dt;
                    button3.Enabled = true;
                    //button5.Enabled = true; 
                    MessageBox.Show("Đã load  "+ dt.Rows.Count + "dòng");
                }
                else
                {
                    MessageBox.Show("Không có kết quả nào");
                }
                button8.Enabled = true;

            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi kết nối ! " + ex.ToString());
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
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

                        //dataGridView1.Rows[rowindex].Cells[columnindex].Value.ToString();

                        var selectedRows = dataGridView1.SelectedRows
                       .OfType<DataGridViewRow>()
                       .Where(row => !row.IsNewRow)
                       .ToArray();


                        progressBar1.Minimum = 0; //Đặt giá trị nhỏ nhất cho ProgressBar
                        progressBar1.Maximum = dataGridView1.SelectedRows.Count; //Đặt giá trị lớn nhất cho ProgressBar
                        int startprogress = 0;


                        foreach (var row in selectedRows)
                        {
                            string ycsxno = row.Cells[3].Value.ToString();
                            dt = pro.getFullInfo(ycsxno);
                            if (file != "")
                            {
                                ExcelFactory.editFileExcelQLSX(file, dt, checkBox1, saveycsxpath);
                                startprogress = startprogress + 1;
                                label4.Text = "Progress: " + startprogress + "/" + dataGridView1.SelectedRows.Count;
                                progressBar1.Value = startprogress;


                            }
                        }
                        MessageBox.Show("Export Yêu cầu và thêm chỉ thị hoàn thành !");
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

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
           

        }

        private void button11_Click(object sender, EventArgs e)
        {
            Form3 frm3 = new Form3();
            frm3.loginIDfrm3 = LoginID;
            frm3.Show();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            try
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("G_NAME");
                dt.Columns.Add("DRAW_PDF_FILENAME");
                listchuabanve = new List<string> { };
                string Dir = System.IO.Directory.GetCurrentDirectory();
                var selectedRows = dataGridView1.SelectedRows
                .OfType<DataGridViewRow>()
                .Where(row => !row.IsNewRow)
                .ToArray();
                foreach (var row in selectedRows)
                {
                    string gname = row.Cells[0].Value.ToString().Substring(0, 11);
                    string gcode = row.Cells[7].Value.ToString().Substring(7, 1);
                    string drawpath = Dir + "\\BANVE\\" + gname + "_" + gcode + ".pdf";
                    //MessageBox.Show(drawpath);
                    if (File.Exists(drawpath))
                    {

                    }
                    else
                    {
                        //MessageBox.Show("Không có bản vẽ : " + gname);
                        dt.Rows.Add(new object[] { row.Cells[0].Value.ToString(), gname + "_" + gcode + ".pdf" });
                    }
                }
                dataGridView1.DataSource = dt;
                MessageBox.Show("Đã check xong bản vẽ !, đây là list code chưa có bản vẽ, chuyển vào thư mục BANVE nhé");
            }

            catch(Exception ex)
            {
                MessageBox.Show("Hãy search YCSX để có thể check bản vẽ !" + ex.ToString());
            }
           

        }

        public string returnLotArray(DataGridView dtgv1)
        {
            string lotArray = "'";
            string finalarray = "'";
            if (dtgv1 == null || dtgv1.Rows.Count == 0)
            {
                MessageBox.Show("Import data vào trước  !");
            }
            else
            {
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();
                foreach (DataGridViewRow row in dtgv1.Rows)
                {
                    if (!row.IsNewRow)
                    {
                        lotArray = lotArray + Convert.ToString(row.Cells[0].Value) + "','";
                    }
                }
                 finalarray = "("+lotArray.Substring(0, lotArray.Length - 2)+")";                
            }
            return finalarray;
        }

        public string returnCode(DataGridView dtgv1)
        {
            string lotArray = "(M100.G_NAME LIKE '";
            string finalarray = "'";
            if (dtgv1 == null || dtgv1.Rows.Count == 0)
            {
                MessageBox.Show("Import data vào trước  !");
            }
            else
            {
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();
                foreach (DataGridViewRow row in dtgv1.Rows)
                {
                    if (!row.IsNewRow)
                    {
                        lotArray = lotArray + Convert.ToString(row.Cells[0].Value) + "%') OR (M100.G_NAME LIKE '";
                    }
                }

                finalarray =  lotArray.Substring(0, lotArray.Length - 23) ;
               
            }
            return finalarray;
        }

        public string nextOutNo(string lastoutno)
        {
            string nextout = "";
            int currentout = int.Parse(lastoutno);
            
            currentout++;
            if(currentout>999)
            {
                currentout = 1;
            }

            if(currentout<10)
            {
                nextout= "00" + currentout;
            }
            else if (currentout < 100)
            {
                nextout= "0" + currentout;
            }
            else if (currentout < 1000)
            {
                nextout= "" + currentout;
            }
            return nextout;


        }
        private void button14_Click(object sender, EventArgs e)
        {
            try
            {
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();
                DataTable dt2 = new DataTable();
                //dataGridView1.Rows[rowindex].Cells[columnindex].Value.ToString();
                var selectedRows = dataGridView1.SelectedRows
                .OfType<DataGridViewRow>()
                .Where(row => !row.IsNewRow)
                .ToArray();

                progressBar1.Minimum = 0; //Đặt giá trị nhỏ nhất cho ProgressBar
                progressBar1.Maximum = dataGridView1.SelectedRows.Count; //Đặt giá trị lớn nhất cho ProgressBar
                int startprogress = 0;

                foreach (var row in selectedRows)
                {
                    string ycsxno = row.Cells[3].Value.ToString();
                    dt = pro.getFullInfo(ycsxno);
                    string today = STYMD(DateTime.Today.Year, DateTime.Today.Month, DateTime.Today.Day);
                    dt2 = pro.getLastOutNo(today);
                    dataGridView1.DataSource = dt2;
                    string lastoutno = dt2.Rows[0]["OUT_NO"].ToString();
                    string nextoutno = nextOutNo(lastoutno);
                    //MessageBox.Show(today+"-" + nextoutno);
                    dataGridView1.DataSource = dt;

                    string CODE_03 = dt.Rows[0]["CODE_03"].ToString();
                    string CODE_50 = dt.Rows[0]["CODE_50"].ToString();
                    string CODE_52 = "01"; //line
                    string DEPT_CD = textBox1.Text; //bo phan lay lieu
                    string PROD_REQUEST_DATE = dt.Rows[0]["PROD_REQUEST_DATE"].ToString();
                    string PROD_REQUEST_NO = dt.Rows[0]["PROD_REQUEST_NO"].ToString();
                    
                    startprogress = startprogress + 1;
                    label4.Text = "Progress: " + startprogress + "/" + dataGridView1.SelectedRows.Count;
                    progressBar1.Value = startprogress;
                }
                MessageBox.Show("Thêm yêu cầu xuất kho hàng loạt hoàn thành !");
                progressBar1.Value = 0;                                           

            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.ToString());
            }
        }

        public static void task1_func()
        {
            Thread.Sleep(2000);
            MessageBox.Show("Xong task1");
        }
        public static void task2_func()
        {
            Thread.Sleep(4000);
            MessageBox.Show("Xong task2");
        }
        private async static void SongSong()
        {
            var task1 = Task.Factory.StartNew(task1_func);
            var task2 = Task.Factory.StartNew(task2_func);
            var task3 = Task.Factory.StartNew(task1_func);
            
            MessageBox.Show("Start song song");
            Task.WaitAll(task1, task2,task3);
            MessageBox.Show("Finish tat ca song song");

        }
        private void button15_Click(object sender, EventArgs e)
        {
            //ProductBLL pro = new ProductBLL();
            //DataTable dt = new DataTable();
            //dt = pro.testQuery(textBox1.Text);
            //dataGridView1.DataSource = dt;
            //MessageBox.Show("Đã load : " + dt.Rows.Count.ToString() + " dòng");
            //button3.Enabled = true;
            SongSong();
        }

        private void button16_Click(object sender, EventArgs e)
        {            
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            //MessageBox.Show(returnCode(dataGridView1));
            dt = pro.checkLastYCSX(returnCode(dataGridView1));   
            //dt = pro.testQuery(textBox1.Text);
            dataGridView1.DataSource = dt;
            MessageBox.Show("Đã load : " + dt.Rows.Count.ToString() + " dòng");
        }

        private void button17_Click(object sender, EventArgs e)
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
                                DataTable dt = new DataTable();
                                pro.InsertPO(CTR_CD, CUST_CD, EMPL_NO, G_CODE, PO_NO, PO_QTY, PO_DATE, RD_DATE, PROD_PRICE);                               

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

        private void button18_Click(object sender, EventArgs e)
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
                                DataTable dt = new DataTable();
                                //pro.InsertPO(CTR_CD, CUST_CD, EMPL_NO, G_CODE, PO_NO, PO_QTY, PO_DATE, RD_DATE, PROD_PRICE);
                                MessageBox.Show(CTR_CD + CUST_CD + EMPL_NO + G_CODE + PO_NO + PO_QTY + PO_DATE + RD_DATE + PROD_PRICE);

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

        private void button20_Click(object sender, EventArgs e)
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
                            string CUST_CD, EMPL_NO, G_CODE, PO_NO, DELIVERY_QTY, DELIVERY_DATE, NOCANCEL;
                            string CTR_CD = "002";

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
                                DataTable dt = new DataTable();
                                pro.InsertInvoice(CTR_CD, CUST_CD, EMPL_NO, G_CODE, PO_NO, DELIVERY_QTY, DELIVERY_DATE, NOCANCEL);

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


        public void traPO()
        {
            try
            {
                bom_flag = 0;
                fcst_flag = 0;
                invoice_flag = 0;
                khgh_flag = 0;
                po_flag = 1;
                ycsx_flag = 0;

                button4.Enabled = true;
                string fromdate, todate;
                fromdate = STYMD(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day);
                todate = STYMD(dateTimePicker2.Value.Year, dateTimePicker2.Value.Month, dateTimePicker2.Value.Day);
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();
                DataTable dttong = new DataTable();
                dt = pro.traPO(generate_condition());
                dttong = pro.traPOTotal(generate_condition());
                dataGridView1.DataSource = dt;
                
                if (dt.Rows.Count > 0)
                {
                    this.dataGridView1.DataSource = null;
                    this.dataGridView1.Columns.Clear();

                    dataGridView1.DataSource = dt;
                    setRowNumber(dataGridView1);
                    button3.Enabled = true;
                    //button5.Enabled = true; 

                    textBox2.Text = dttong.Rows[0]["PO_QTY"].ToString();
                    textBox2.Text = string.Format("{0:#,##0 EA}", double.Parse(textBox2.Text));
                    textBox4.Text = dttong.Rows[0]["TOTAL_DELIVERED"].ToString();
                    textBox4.Text = string.Format("{0:#,##0 EA}", double.Parse(textBox4.Text));
                    textBox6.Text = dttong.Rows[0]["PO_BALANCE"].ToString();
                    textBox6.Text = string.Format("{0:#,##0 EA}", double.Parse(textBox6.Text));

                    textBox3.Text = dttong.Rows[0]["PO_AMOUNT"].ToString();
                    textBox3.Text = string.Format("{0:#,##0.00 $}", double.Parse(textBox3.Text));
                    textBox5.Text = dttong.Rows[0]["DELIVERED_AMOUNT"].ToString();
                    textBox5.Text = string.Format("{0:#,##0.00 $}", double.Parse(textBox5.Text));
                    textBox7.Text = dttong.Rows[0]["BALANCE_AMOUNT"].ToString();
                    textBox7.Text = string.Format("{0:#,##0.00 $}", double.Parse(textBox7.Text));

                    textBox2.Font = new Font("Microsoft Sans Serif", 15.0f);
                    textBox4.Font = new Font("Microsoft Sans Serif", 15.0f);
                    textBox6.Font = new Font("Microsoft Sans Serif", 15.0f);
                    textBox3.Font = new Font("Microsoft Sans Serif", 15.0f);
                    textBox5.Font = new Font("Microsoft Sans Serif", 15.0f);
                    textBox7.Font = new Font("Microsoft Sans Serif", 15.0f);

                    formatDataGridViewtraPO1(dataGridView1);

                    MessageBox.Show("Đã load  " + dt.Rows.Count + "dòng");
                    po_flag = 1;
                    invoice_flag = 0;
                    khgh_flag = 0;
                    fcst_flag = 0;
                    ycsx_flag = 0;
                }
                else
                {
                    MessageBox.Show("Không có kết quả nào");
                }                
                button8.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi kết nối ! " + ex.ToString());
            }
        }

        public void traPO_Async(DoWorkEventArgs e)
        {
            try
            {               

                button4.Enabled = true;
                string fromdate, todate;
                fromdate = STYMD(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day);
                todate = STYMD(dateTimePicker2.Value.Year, dateTimePicker2.Value.Month, dateTimePicker2.Value.Day);
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();
                DataTable dttong = new DataTable();
                dtgv1_data = pro.traPO(generate_condition());
                dtgv1_total_data = pro.traPOTotal(generate_condition());
               
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi kết nối ! " + ex.ToString());
            }
        }

        public void formatDataGridViewtraPO1(DataGridView dataGridView1)
        {
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;

            dataGridView1.Columns["PO_QTY"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["TOTAL_DELIVERED"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["PO_BALANCE"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["DELIVERED_AMOUNT"].DefaultCellStyle.Format = "c";
            dataGridView1.Columns["BALANCE_AMOUNT"].DefaultCellStyle.Format = "c";
            dataGridView1.Columns["PO_AMOUNT"].DefaultCellStyle.Format = "c";

            dataGridView1.Columns["TON_KIEM"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["BTP"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["TP"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["BLOCK_QTY"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["GRAND_TOTAL_STOCK"].DefaultCellStyle.Format = "#,0";



            dataGridView1.Columns["PO_QTY"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["PO_QTY"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["PO_QTY"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["PO_BALANCE"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["PO_BALANCE"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["PO_BALANCE"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["TOTAL_DELIVERED"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["TOTAL_DELIVERED"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["TOTAL_DELIVERED"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["DELIVERED_AMOUNT"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["DELIVERED_AMOUNT"].DefaultCellStyle.BackColor = Color.Black;
            dataGridView1.Columns["DELIVERED_AMOUNT"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["BALANCE_AMOUNT"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["BALANCE_AMOUNT"].DefaultCellStyle.BackColor = Color.Black;
            dataGridView1.Columns["BALANCE_AMOUNT"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["PO_AMOUNT"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["PO_AMOUNT"].DefaultCellStyle.BackColor = Color.Black;
            dataGridView1.Columns["PO_AMOUNT"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);


            dataGridView1.Columns["CUST_NAME_KD"].DefaultCellStyle.BackColor = Color.Aqua;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.BackColor = Color.Brown;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.ForeColor = Color.Yellow;

            dataGridView1.Columns["GRAND_TOTAL_STOCK"].DefaultCellStyle.ForeColor = Color.Black;
            dataGridView1.Columns["GRAND_TOTAL_STOCK"].DefaultCellStyle.BackColor = Color.GreenYellow;
            dataGridView1.Columns["GRAND_TOTAL_STOCK"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);


            dataGridView1.Columns["TON_KIEM"].DefaultCellStyle.ForeColor = Color.Blue;
            dataGridView1.Columns["TON_KIEM"].DefaultCellStyle.BackColor = Color.Aqua;
            dataGridView1.Columns["TON_KIEM"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Regular);

            dataGridView1.Columns["BTP"].DefaultCellStyle.ForeColor = Color.Blue;
            dataGridView1.Columns["BTP"].DefaultCellStyle.BackColor = Color.Aqua;
            dataGridView1.Columns["BTP"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Regular);

            dataGridView1.Columns["TP"].DefaultCellStyle.ForeColor = Color.Blue;
            dataGridView1.Columns["TP"].DefaultCellStyle.BackColor = Color.Aqua;
            dataGridView1.Columns["TP"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Regular);

            dataGridView1.Columns["BLOCK_QTY"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["BLOCK_QTY"].DefaultCellStyle.BackColor = Color.Red;
            dataGridView1.Columns["BLOCK_QTY"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Regular);


        }

        public void formatDataGridViewtraInvoice1(DataGridView dataGridView1)
        {
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;

            dataGridView1.Columns["DELIVERY_QTY"].DefaultCellStyle.Format = "#,0";                     
            dataGridView1.Columns["DELIVERED_AMOUNT"].DefaultCellStyle.Format = "c";
            dataGridView1.Columns["DELIVERED_AMOUNT"].DefaultCellStyle.Format = "c";

            
            dataGridView1.Columns["DELIVERY_QTY"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["DELIVERY_QTY"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["DELIVERY_QTY"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["DELIVERED_AMOUNT"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["DELIVERED_AMOUNT"].DefaultCellStyle.BackColor = Color.Black;
            dataGridView1.Columns["DELIVERED_AMOUNT"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);


            dataGridView1.Columns["CUST_NAME_KD"].DefaultCellStyle.BackColor = Color.Aqua;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.BackColor = Color.Brown;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.ForeColor = Color.Yellow;


        }

        public void formatDataGridViewtraFCST1(DataGridView dataGridView1)
        {
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;

            dataGridView1.Columns["W1"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W2"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W3"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W4"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W5"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W6"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W7"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W8"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W9"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W10"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W11"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W12"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W13"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W14"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W15"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W16"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W17"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W18"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W19"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W20"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W21"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W22"].DefaultCellStyle.Format = "#,0";

            dataGridView1.Columns["W1A"].DefaultCellStyle.Format = "c";
            dataGridView1.Columns["W2A"].DefaultCellStyle.Format = "c";
            dataGridView1.Columns["W3A"].DefaultCellStyle.Format = "c";
            dataGridView1.Columns["W4A"].DefaultCellStyle.Format = "c";
            dataGridView1.Columns["W5A"].DefaultCellStyle.Format = "c";
            dataGridView1.Columns["W6A"].DefaultCellStyle.Format = "c";
            dataGridView1.Columns["W7A"].DefaultCellStyle.Format = "c";
            dataGridView1.Columns["W8A"].DefaultCellStyle.Format = "c";
            dataGridView1.Columns["W9A"].DefaultCellStyle.Format = "c";
            dataGridView1.Columns["W10A"].DefaultCellStyle.Format = "c";
            dataGridView1.Columns["W11A"].DefaultCellStyle.Format = "c";
            dataGridView1.Columns["W12A"].DefaultCellStyle.Format = "c";
            dataGridView1.Columns["W13A"].DefaultCellStyle.Format = "c";
            dataGridView1.Columns["W14A"].DefaultCellStyle.Format = "c";
            dataGridView1.Columns["W15A"].DefaultCellStyle.Format = "c";
            dataGridView1.Columns["W16A"].DefaultCellStyle.Format = "c";
            dataGridView1.Columns["W17A"].DefaultCellStyle.Format = "c";
            dataGridView1.Columns["W18A"].DefaultCellStyle.Format = "c";
            dataGridView1.Columns["W19A"].DefaultCellStyle.Format = "c";
            dataGridView1.Columns["W20A"].DefaultCellStyle.Format = "c";
            dataGridView1.Columns["W21A"].DefaultCellStyle.Format = "c";
            dataGridView1.Columns["W22A"].DefaultCellStyle.Format = "c";

            dataGridView1.Columns["CUST_NAME_KD"].DefaultCellStyle.BackColor = Color.Aqua;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.BackColor = Color.Brown;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.ForeColor = Color.Yellow;

        }

        public void formatDataGridViewtraPlan1(DataGridView dataGridView1)
        {
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;

            dataGridView1.Columns["D1"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["D2"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["D3"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["D4"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["D5"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["D6"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["D7"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["D8"].DefaultCellStyle.Format = "#,0";

            dataGridView1.Columns["CUST_NAME_KD"].DefaultCellStyle.BackColor = Color.Aqua;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.BackColor = Color.Brown;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.ForeColor = Color.Yellow;

        }



        public void formatDataGridViewtraPO(DataGridView dataGridView1)
        {
            try
            {
                dataGridView1.EnableHeadersVisualStyles = false;
                dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
                dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
                dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
                dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
                dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;

                dataGridView1.Columns["PO_QTY"].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns["TOTAL_DELIVERED"].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns["PO_BALANCE"].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns["DELIVERED_AMOUNT"].DefaultCellStyle.Format = "c";
                dataGridView1.Columns["BALANCE_AMOUNT"].DefaultCellStyle.Format = "c";
                dataGridView1.Columns["PO_AMOUNT"].DefaultCellStyle.Format = "c";

                dataGridView1.Columns["D1"].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns["D2"].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns["D3"].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns["D4"].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns["D5"].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns["D6"].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns["D7"].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns["D8"].DefaultCellStyle.Format = "#,0";



                dataGridView1.Columns["PO_QTY"].DefaultCellStyle.ForeColor = Color.White;
                dataGridView1.Columns["PO_QTY"].DefaultCellStyle.BackColor = Color.Gray;
                dataGridView1.Columns["PO_QTY"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

                dataGridView1.Columns["PO_BALANCE"].DefaultCellStyle.ForeColor = Color.White;
                dataGridView1.Columns["PO_BALANCE"].DefaultCellStyle.BackColor = Color.Gray;
                dataGridView1.Columns["PO_BALANCE"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

                dataGridView1.Columns["TOTAL_DELIVERED"].DefaultCellStyle.ForeColor = Color.White;
                dataGridView1.Columns["TOTAL_DELIVERED"].DefaultCellStyle.BackColor = Color.Gray;
                dataGridView1.Columns["TOTAL_DELIVERED"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

                dataGridView1.Columns["DELIVERED_AMOUNT"].DefaultCellStyle.ForeColor = Color.Yellow;
                dataGridView1.Columns["DELIVERED_AMOUNT"].DefaultCellStyle.BackColor = Color.Black;
                dataGridView1.Columns["DELIVERED_AMOUNT"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

                dataGridView1.Columns["BALANCE_AMOUNT"].DefaultCellStyle.ForeColor = Color.Yellow;
                dataGridView1.Columns["BALANCE_AMOUNT"].DefaultCellStyle.BackColor = Color.Black;
                dataGridView1.Columns["BALANCE_AMOUNT"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

                dataGridView1.Columns["PO_AMOUNT"].DefaultCellStyle.ForeColor = Color.Yellow;
                dataGridView1.Columns["PO_AMOUNT"].DefaultCellStyle.BackColor = Color.Black;
                dataGridView1.Columns["PO_AMOUNT"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);


                dataGridView1.Columns["CUST_NAME_KD"].DefaultCellStyle.BackColor = Color.Aqua;
                dataGridView1.Columns["G_NAME"].DefaultCellStyle.BackColor = Color.Brown;
                dataGridView1.Columns["G_NAME"].DefaultCellStyle.ForeColor = Color.Yellow;


            }
            catch(Exception ex)
            {

            }





        }

        public void formatDataGridViewChoKiem(DataGridView dataGridView1)
        {
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;

            dataGridView1.Columns["INSPECT_BALANCE_QTY"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["WAIT_CS_QTY"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["WAIT_SORTING_RMA"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["TOTAL_WAIT"].DefaultCellStyle.Format = "#,0";
          

            dataGridView1.Columns["G_CODE"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["G_CODE"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["G_CODE"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["INSPECT_BALANCE_QTY"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["INSPECT_BALANCE_QTY"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["INSPECT_BALANCE_QTY"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["WAIT_CS_QTY"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["WAIT_CS_QTY"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["WAIT_CS_QTY"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["WAIT_SORTING_RMA"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["WAIT_SORTING_RMA"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["WAIT_SORTING_RMA"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["TOTAL_WAIT"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["TOTAL_WAIT"].DefaultCellStyle.BackColor = Color.Black;
            dataGridView1.Columns["TOTAL_WAIT"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["G_NAME"].DefaultCellStyle.BackColor = Color.Brown;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.ForeColor = Color.Yellow;

        }

        public void formatDataGridViewChoKiem2(DataGridView dataGridView1)
        {
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;

            dataGridView1.Columns["INSPECT_BALANCE_QTY"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["WAIT_CS_QTY"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["WAIT_SORTING_RMA"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["TOTAL_WAIT"].DefaultCellStyle.Format = "#,0";


            dataGridView1.Columns["G_CODE"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["G_CODE"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["G_CODE"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["INSPECT_BALANCE_QTY"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["INSPECT_BALANCE_QTY"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["INSPECT_BALANCE_QTY"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["WAIT_CS_QTY"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["WAIT_CS_QTY"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["WAIT_CS_QTY"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["WAIT_SORTING_RMA"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["WAIT_SORTING_RMA"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["WAIT_SORTING_RMA"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["TOTAL_WAIT"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["TOTAL_WAIT"].DefaultCellStyle.BackColor = Color.Black;
            dataGridView1.Columns["TOTAL_WAIT"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["G_NAME"].DefaultCellStyle.BackColor = Color.Brown;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.ForeColor = Color.Yellow;

        }

        public void formatDataGridViewBTP(DataGridView dataGridView1)
        {
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;

            dataGridView1.Columns["BTP_QTY_EA"].DefaultCellStyle.Format = "#,0";       


            dataGridView1.Columns["G_CODE"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["G_CODE"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["G_CODE"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["G_NAME"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["BTP_QTY_EA"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["BTP_QTY_EA"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["BTP_QTY_EA"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);
         

            dataGridView1.Columns["G_NAME"].DefaultCellStyle.BackColor = Color.Brown;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.ForeColor = Color.Yellow;

        }



        public void tracuu_KD(DoWorkEventArgs e)
        {
            
            if (po_flag ==1)
            {
                pictureBox1.Show();
                traPO_Async(e);                
            }
            else if (invoice_flag == 1)
            {
                pictureBox1.Show();
                traInvoice_Async(e);
            }
            else if (fcst_flag == 1)
            {
                pictureBox1.Show();
                traFCST_Async(e);
            }
            else if (khgh_flag == 1)
            {
                pictureBox1.Show();
                traPLAN_Async(e);
            }
            else if (ycsx_flag == 1)
            {

            }
            else if (bom_flag == 1)
            {
                pictureBox1.Show();
                traBOM_Async(e);
            }
            else if (check_po_flag == 1)
            {
                pictureBox1.Show();
                checkPO_Async(e);
            }
            else if (check_invoice_flag == 1)
            {
                pictureBox1.Show();
                checkInvoice_Async(e);
            }
            else if (check_fcst_flag == 1)
            {
                pictureBox1.Show();
                checkFCST_Async(e);
            }
            else if (check_plan_flag == 1)
            {
                pictureBox1.Show();
                checkPLAN_Async(e);
            }
            else if (up_po_flag == 1)
            {
                pictureBox1.Show();
                upPOhangloat_Async(e);
            }
            else if (up_invoice_flag == 1)
            {
                pictureBox1.Show();
                upInvoicehangloat_Async(e);
            }
            else if (up_fcst_flag == 1)
            {
                pictureBox1.Show();
                uploadFCST_Async(e);
            }
            else if (up_plan_flag == 1)
            {
                pictureBox1.Show();
                upPLAN_Async(e);
            }
        }

      
        private void button21_Click(object sender, EventArgs e)
        {
            dataGridView1.Show();
            traPO();          

        }

        private void button25_Click(object sender, EventArgs e)
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
                            string   G_CODE, G_CODE_KD, PROD_TYPE, PROD_MODEL, PROD_PROJECT, PROD_SIZE, PROD_MAIN_MATERIAL;

                            G_CODE = Convert.ToString(row.Cells[0].Value);
                            G_CODE_KD = Convert.ToString(row.Cells[1].Value);
                            PROD_TYPE = Convert.ToString(row.Cells[2].Value);
                            PROD_MODEL = Convert.ToString(row.Cells[3].Value);
                            PROD_PROJECT = Convert.ToString(row.Cells[4].Value);
                            PROD_SIZE = Convert.ToString(row.Cells[5].Value);
                            PROD_MAIN_MATERIAL = Convert.ToString(row.Cells[6].Value);
                            


                           /* if ((G_CODE == "") || (G_CODE_KD == "") || (PROD_TYPE == "") || (PROD_MODEL == "") || (PROD_PROJECT == "") || (PROD_SIZE == "") || (PROD_MAIN_MATERIAL == ""))
                            {
                                MessageBox.Show("Thông tin của sản phẩm: " + G_CODE + "có thông tin trống, sẽ ko thêm PO này !");
                            }
                            */
                                ProductBLL pro = new ProductBLL();
                                DataTable dt = new DataTable();
                                pro.updateInfo(G_CODE, G_CODE_KD, PROD_TYPE, PROD_MODEL, PROD_PROJECT,  PROD_MAIN_MATERIAL);
                            

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

        private void groupBox5_Enter(object sender, EventArgs e)
        {

        }

        private void button26_Click(object sender, EventArgs e)
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
                            CUST_NAME_KD = Convert.ToString(row.Cells[1].Value);
                           

                           
                            ProductBLL pro = new ProductBLL();
                            DataTable dt = new DataTable();
                            pro.updateCustomer(CUST_CD, CUST_NAME_KD);


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

        private void tạoPOMớiToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void traPOToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void yêuCầuSảnXuấtCủaTôiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                button4.Enabled = true;
                string fromdate, todate;
                fromdate = STYMD(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day);
                todate = STYMD(dateTimePicker2.Value.Year, dateTimePicker2.Value.Month, dateTimePicker2.Value.Day);
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();
                dt = pro.traYCSXPIC(LoginID, textBox1.Text, fromdate, todate);
                if (dt.Rows.Count > 0)
                {
                    this.dataGridView1.DataSource = null;
                    this.dataGridView1.Columns.Clear();
                    dataGridView1.DataSource = dt;
                    button3.Enabled = true;
                    //button5.Enabled = true; 
                    MessageBox.Show("Đã load  " + dt.Rows.Count + "dòng");
                }
                else
                {
                    MessageBox.Show("Không có kết quả nào");
                }
                button8.Enabled = true;

            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi kết nối ! " + ex.ToString());
            }
        }

        private void tấtCảYêuCầuSảnXuấtToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                button4.Enabled = true;
                string fromdate, todate;
                fromdate = STYMD(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day);
                todate = STYMD(dateTimePicker2.Value.Year, dateTimePicker2.Value.Month, dateTimePicker2.Value.Day);
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();
                dt = pro.traYCSX(textBox1.Text, fromdate, todate);
                if (dt.Rows.Count > 0)
                {
                    this.dataGridView1.DataSource = null;
                    this.dataGridView1.Columns.Clear();
                    dataGridView1.DataSource = dt;
                    button3.Enabled = true;
                    MessageBox.Show("Đã load " + dt.Rows.Count + " dòng");
                    //button5.Enabled = true;
                }
                else
                {
                    MessageBox.Show("Không có kết quả nào");
                }
                button8.Enabled = true;

            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi kết nối ! " + ex.ToString());
            }
        }

        private void tạo1YCSXMớiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 frm3 = new Form3();
            frm3.loginIDfrm3 = LoginID;
            frm3.Show();
        }

        private void uploadYCSXHàngLoạtToolStripMenuItem_Click(object sender, EventArgs e)
        {

            try
            {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.ShowDialog();
                string file = ofd.FileName;

                if (file != "")
                {
                    dsNV = ExcelFactory.readFromExcelFile(file);
                    dataGridView1.DataSource = dsNV;
                    button4.Enabled = true;
                    button3.Enabled = true;
                    button16.Enabled = true;
                    /*
                    dataGridView1.Columns[0].HeaderText = "Mã Nhân Viên";
                    dataGridView1.Columns[1].HeaderText = "Tên Nhân Viên";
                    dataGridView1.Columns[2].HeaderText = "Tuổi Nhân Viên";
                    */

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
                            //PROD_REQUEST_DATE = Convert.ToString(row.Cells[0].Value);
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
                                    if (lastycsxno.Substring(0, 3) != CreateHeader2())
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

                                    string PROD_REQUEST_NO = CreateHeader2() + lastycsxno;

                                    int check_riv = pro.checkRIV_NO(G_CODE, RIV_NO);

                                    string checkUSEYN = pro.checkM100UseYN(G_CODE);
                                    /*
                                    if (check_riv != 1)
                                    {
                                        MessageBox.Show("Code " + G_CODE + " không tồn tại REVISION trong bảng BOM, check lại REVISION hoặc liên hệ RND");
                                    }
                                    else if (checkUSEYN == "N")
                                    {
                                        MessageBox.Show("Code " + G_CODE + " đã bị khóa, có thể ver này không còn được sử dụng");
                                    }
                                    else
                                    {
                                        pro.InsertYCSX(CTR_CD, PROD_REQUEST_DATE, PROD_REQUEST_NO, CODE_50, CODE_03, CODE_55, G_CODE, RIV_NO, PROD_REQUEST_QTY, CUST_CD, EMPL_NO, REMK, EMPL_NO, EMPL_NO, DELIVERY_DT);
                                        pro.writeHistory("002", LoginID, "YCSX TABLE", "THEM", "THEM YCSX", "0");
                                    }
                                    */

                                    if (checkUSEYN == "N")
                                    {
                                        MessageBox.Show("Code " + G_CODE + " đã bị khóa, có thể ver này không còn được sử dụng");
                                    }
                                    else
                                    {
                                        //pro.InsertYCSX(CTR_CD, PROD_REQUEST_DATE, PROD_REQUEST_NO, CODE_50, CODE_03, CODE_55, G_CODE, "A", PROD_REQUEST_QTY, CUST_CD, EMPL_NO, REMK, EMPL_NO, EMPL_NO, DELIVERY_DT);
                                        pro.writeHistory(CTR_CD, EMPL_NO, "YCSX TABLE", "THEM", "THEM YCSX", "0");
                                    }

                                }
                                else
                                {
                                    MessageBox.Show("Loi");
                                }
                            }


                            /*
                            for (int i = 0; i < dataGridView1.Columns.Count; i++)
                            {
                                String header = dataGridView1.Columns[i].HeaderText;
                                String cellText = Convert.ToString(row.Cells[i].Value);

                                //MessageBox.Show(cellText);
                            }
                            */


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

        private void checkBảnVẽToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("G_NAME");
                dt.Columns.Add("DRAW_PDF_FILENAME");
                listchuabanve = new List<string> { };
                string Dir = System.IO.Directory.GetCurrentDirectory();
                var selectedRows = dataGridView1.SelectedRows
                .OfType<DataGridViewRow>()
                .Where(row => !row.IsNewRow)
                .ToArray();
                foreach (var row in selectedRows)
                {
                    string gname = row.Cells[0].Value.ToString().Substring(0, 11);
                    string gcode = row.Cells[7].Value.ToString().Substring(7, 1);
                    string drawpath = Dir + "\\BANVE\\" + gname + "_" + gcode + ".pdf";
                    //MessageBox.Show(drawpath);
                    if (File.Exists(drawpath))
                    {

                    }
                    else
                    {
                        //MessageBox.Show("Không có bản vẽ : " + gname);
                        dt.Rows.Add(new object[] { row.Cells[0].Value.ToString(), gname + "_" + gcode + ".pdf" });
                    }
                }
                dataGridView1.DataSource = dt;
                MessageBox.Show("Đã check xong bản vẽ !, đây là list code chưa có bản vẽ, chuyển vào thư mục BANVE nhé");
            }

            catch (Exception ex)
            {
                MessageBox.Show("Hãy search YCSX để có thể check bản vẽ !" + ex.ToString());
            }
        }

        private void xuấtFileYCSXOnlyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            checkBox1.Checked = false;
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

                        var selectedRows = dataGridView1.SelectedRows
                       .OfType<DataGridViewRow>()
                       .Where(row => !row.IsNewRow)
                       .ToArray();

                        progressBar1.Minimum = 0; //Đặt giá trị nhỏ nhất cho ProgressBar
                        progressBar1.Maximum = dataGridView1.SelectedRows.Count; //Đặt giá trị lớn nhất cho ProgressBar
                        int startprogress = 0;
                        /*
                        for(int kk= 0; kk< selectedRows.Length; kk ++)
                        {
                            string ycsxno = selectedRows[kk].Cells[3].Value.ToString();
                            dt = pro.getFullInfo(ycsxno);
                            if (file != "")
                            {
                                ExcelFactory.editFileExcel(file, dt, checkBox1, saveycsxpath);
                                startprogress = startprogress + 1;
                                label4.Text = "Progress: " + startprogress + "/" + dataGridView1.SelectedRows.Count;
                                progressBar1.Value = startprogress;
                            }

                        }
                        */

                        foreach (var row in selectedRows)
                        {
                            string ycsxno = row.Cells[3].Value.ToString();
                            dt = pro.getFullInfo(ycsxno);
                            if (file != "")
                            {
                                ExcelFactory.editFileExcel(file, dt, checkBox1, saveycsxpath);
                                startprogress = startprogress + 1;
                                label4.Text = "Progress: " + startprogress + "/" + dataGridView1.SelectedRows.Count;
                                progressBar1.Value = startprogress;

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

        private void xuấtFileYCSXVàInYCSXToolStripMenuItem_Click(object sender, EventArgs e)
        {
            checkBox1.Checked = true;
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

                        var selectedRows = dataGridView1.SelectedRows
                       .OfType<DataGridViewRow>()
                       .Where(row => !row.IsNewRow)
                       .ToArray();

                        progressBar1.Minimum = 0; //Đặt giá trị nhỏ nhất cho ProgressBar
                        progressBar1.Maximum = dataGridView1.SelectedRows.Count; //Đặt giá trị lớn nhất cho ProgressBar
                        int startprogress = 0;


                        foreach (var row in selectedRows)
                        {
                            string ycsxno = row.Cells[3].Value.ToString();
                            dt = pro.getFullInfo(ycsxno);
                            if (file != "")
                            {
                                string drawfilename = dt.Rows[0]["G_NAME"].ToString().Substring(0, 11) + "_" + dt.Rows[0]["G_CODE"].ToString().Substring(7, 1) + ".pdf";
                                //MessageBox.Show(drawfilename);
                                string pdffile = Dir + "\\BANVE\\" + drawfilename;
                                //MessageBox.Show(pdffile);
                                if (File.Exists(pdffile))
                                {
                                    printPDF(pdffile);
                                }
                                else
                                {
                                    MessageBox.Show("Không có bản vẽ : " + dt.Rows[0]["G_NAME"].ToString());
                                }

                                ExcelFactory.editFileExcel(file, dt, checkBox1, saveycsxpath);
                                startprogress = startprogress + 1;
                                label4.Text = "Progress: " + startprogress + "/" + dataGridView1.SelectedRows.Count;
                                progressBar1.Value = startprogress;
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

                            DateTime podate = DateTime.Parse(PO_DATE);
                            int check_date = checkDate(podate);

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
                                else if (check_date ==0)
                                {
                                    dataGridView1.Rows[row.Index].Cells[8].Value = "NG - Ngày PO lớn hơn ngày hiện tại";
                                    dataGridView1.Rows[row.Index].Cells[8].Style.BackColor = Color.Red;
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

        public void upPOhangloat_Async(DoWorkEventArgs e)
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

                            DateTime podate = DateTime.Parse(PO_DATE);
                            int check_date = checkDate(podate);

                            if ((CUST_CD == "") || (EMPL_NO == "") || (G_CODE == "") || (PO_NO == "") || (PO_QTY == "") || (PO_DATE == "") || (RD_DATE == "") || (PROD_PRICE == ""))
                            {
                                MessageBox.Show("Thông tin của sản phẩm: " + G_CODE + "có thông tin trống, sẽ ko thêm PO này !");
                            }
                            else
                            {
                                ProductBLL pro = new ProductBLL();
                                string checkUSEYN = pro.checkM100UseYN(G_CODE);

                                if (pro.checkPOExist(CUST_CD, G_CODE, PO_NO) != -1)
                                {
                                    dataGridView1.Rows[row.Index].Cells[8].Value = "NG - Đã tồn tại PO";
                                    dataGridView1.Rows[row.Index].Cells[8].Style.BackColor = Color.Red;

                                }
                                else if (check_date == 0)
                                {
                                    dataGridView1.Rows[row.Index].Cells[8].Value = "NG - Ngày PO lớn hơn ngày hiện tại";
                                    dataGridView1.Rows[row.Index].Cells[8].Style.BackColor = Color.Red;
                                }
                                else if(checkUSEYN == "N")
                                {
                                    dataGridView1.Rows[row.Index].Cells[8].Value = "NG - Ver này đang bị khoá, check lại ver";
                                    dataGridView1.Rows[row.Index].Cells[8].Style.BackColor = Color.Red;
                                }
                                else
                                {
                                    
                                    pro.InsertPO(CTR_CD, CUST_CD, EMPL_NO, G_CODE, PO_NO, PO_QTY, PO_DATE, RD_DATE, PROD_PRICE);
                                    dataGridView1.Rows[row.Index].Cells[8].Value = "OK";
                                    dataGridView1.Rows[row.Index].Cells[8].Style.BackColor = Color.LightGreen;
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

        public void upthongtincodeQLSXhangloat()
        {
            if (dataGridView1 == null || dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("Import data vào trước  !");
            }
            else
            {



                try
                {
                    ProductBLL pro = new ProductBLL();
                    DataTable dt = new DataTable();

                    var selectedRows = dataGridView1.SelectedRows
                   .OfType<DataGridViewRow>()
                   .Where(row => !row.IsNewRow)
                   .ToArray();

                    progressBar1.Minimum = 0; //Đặt giá trị nhỏ nhất cho ProgressBar
                    progressBar1.Maximum = dataGridView1.SelectedRows.Count; //Đặt giá trị lớn nhất cho ProgressBar
                    int startprogress = 0;

                    foreach (var row in selectedRows)
                    {
                        if (!row.IsNewRow)
                        {
                            string G_CODE, FACTORY, EQ1, EQ2, SETTING1, SETTING2, UPH1, UPH2, STEP1, STEP2, NOTE;


                            G_CODE = Convert.ToString(row.Cells[0].Value);
                            FACTORY = Convert.ToString(row.Cells[7].Value);
                            EQ1 = Convert.ToString(row.Cells[8].Value);
                            EQ2 = Convert.ToString(row.Cells[9].Value);
                            SETTING1 = Convert.ToString(row.Cells[10].Value);
                            SETTING2 = Convert.ToString(row.Cells[11].Value);
                            UPH1 = Convert.ToString(row.Cells[12].Value);
                            UPH2 = Convert.ToString(row.Cells[13].Value);
                            STEP1 = Convert.ToString(row.Cells[14].Value);
                            STEP2 = Convert.ToString(row.Cells[15].Value);
                            NOTE = Convert.ToString(row.Cells[16].Value);

                            int checkflag = 0;

                            if (FACTORY != "NM1" && FACTORY != "NM3" && FACTORY != "NM4")
                            {
                                checkflag = 1;
                            }
                           
                            if (checkflag == 1)
                            {
                                MessageBox.Show("Không đc để trống nhà máy, ko thêm code này : " + G_CODE);
                            }
                            else
                            {
                                pro.updateInfoQLSX(G_CODE, FACTORY, EQ1, EQ2, SETTING1, SETTING2, UPH1, UPH2, STEP1, STEP2, NOTE);
                            }



                            /*
                            int checkflag = 0;

                            if (FACTORY != "NM1" || FACTORY != "NM3" || FACTORY != "NM4")
                            {
                                checkflag = 1;
                            }

                            if (EQ1 == "")
                            {
                                checkflag = 2;
                            }

                            if ((EQ1 != "") && (UPH1 == "") && (STEP1 == "") || (EQ2 != "") && (UPH2 == "") && (STEP2 == ""))
                            {
                                checkflag = 3;
                            }

                            if (checkflag == 1)
                            {
                                MessageBox.Show("Không đc để trống nhà máy, ko thêm code này : " + G_CODE);
                            }
                            if (checkflag == 2)
                            {
                                MessageBox.Show("Không được để trống EQ1: " + G_CODE);
                            }
                            if (checkflag == 3)
                            {
                                MessageBox.Show("Có tên máy phải có thời gian setting, UPH, STEP, k đc để trống nếu có tên máy: " + G_CODE);
                            }
                            else
                            {                                
                                pro.updateInfoQLSX(G_CODE, FACTORY, EQ1, EQ2, SETTING1, SETTING2, UPH1, UPH2, STEP1, STEP2, NOTE);
                            }

                            */

                        }
                        startprogress = startprogress + 1;
                        label4.Text = "Progress: " + startprogress + "/" + dataGridView1.SelectedRows.Count;
                        progressBar1.Value = startprogress;
                    }
                    MessageBox.Show("Đã hoàn thành up info code qlsx của những dòng đã được chọn");
                    progressBar1.Value = 0;
                    dataGridView1.ClearSelection();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }                
            }
        }

        private void uploadHàngLoạtToolStripMenuItem_Click(object sender, EventArgs e)
        {
            upPOhangloat();
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
                        ProductBLL pro = new ProductBLL();
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


                            DateTime dlidate = DateTime.Parse(DELIVERY_DATE);
                            int check_date = checkDate(dlidate);
                            DateTime podate = DateTime.Parse(pro.getPODDate(G_CODE, PO_NO)); // get podate and convert to datetime
                            int checkinvoicedatevspodate = checkInvoicevsPODate(dlidate, podate);

                            if ((CUST_CD == "") || (EMPL_NO == "") || (G_CODE == "") || (PO_NO == "") || (DELIVERY_QTY == "") || (DELIVERY_DATE == "") || (NOCANCEL == ""))
                            {
                                MessageBox.Show("Thông tin của sản phẩm: " + G_CODE + "có thông tin trống, sẽ ko thêm PO này !");
                            }
                            else
                            {
                                
                                int po_balance = pro.checkPOBalance(CUST_CD, G_CODE, PO_NO);
                                
                                if(check_date==0)
                                {
                                    dataGridView1.Rows[row.Index].Cells[8].Value = "NG - Ngày Invoice lớn hơn ngày hiện tại";
                                    dataGridView1.Rows[row.Index].Cells[8].Style.BackColor = Color.Red;
                                }
                                else if (checkinvoicedatevspodate == 0)
                                {
                                    dataGridView1.Rows[row.Index].Cells[8].Value = "NG - Ngày Invoice nhỏ hơn ngày PO";
                                    dataGridView1.Rows[row.Index].Cells[8].Style.BackColor = Color.Red;
                                }
                                else if(int.Parse(DELIVERY_QTY) > po_balance)
                                {
                                    dataGridView1.Rows[row.Index].Cells[8].Value = "NG - Giao hàng nhiều hơn PO";
                                    dataGridView1.Rows[row.Index].Cells[8].Style.BackColor = Color.Red;
                                }
                                else if (int.Parse(DELIVERY_QTY) <= po_balance)
                                {
                                    pro.InsertInvoice(CTR_CD, CUST_CD, EMPL_NO, G_CODE, PO_NO, DELIVERY_QTY, DELIVERY_DATE, NOCANCEL);
                                    dataGridView1.Rows[row.Index].Cells[8].Value = "OK";
                                    dataGridView1.Rows[row.Index].Cells[8].Style.BackColor = Color.Green;
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

        public void upInvoicehangloat_Async(DoWorkEventArgs e)
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
                        ProductBLL pro = new ProductBLL();
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


                            DateTime dlidate = DateTime.Parse(DELIVERY_DATE);
                            int check_date = checkDate(dlidate);
                            DateTime podate = DateTime.Parse(pro.getPODDate(G_CODE, PO_NO)); // get podate and convert to datetime
                            int checkinvoicedatevspodate = checkInvoicevsPODate(dlidate, podate);

                            if ((CUST_CD == "") || (EMPL_NO == "") || (G_CODE == "") || (PO_NO == "") || (DELIVERY_QTY == "") || (DELIVERY_DATE == "") || (NOCANCEL == ""))
                            {
                                MessageBox.Show("Thông tin của sản phẩm: " + G_CODE + "có thông tin trống, sẽ ko thêm PO này !");
                            }
                            else
                            {

                                int po_balance = pro.checkPOBalance(CUST_CD, G_CODE, PO_NO);

                                if (check_date == 0)
                                {
                                    dataGridView1.Rows[row.Index].Cells[8].Value = "NG - Ngày Invoice lớn hơn ngày hiện tại";
                                    dataGridView1.Rows[row.Index].Cells[8].Style.BackColor = Color.Red;
                                }
                                else if (checkinvoicedatevspodate == 0)
                                {
                                    dataGridView1.Rows[row.Index].Cells[8].Value = "NG - Ngày Invoice nhỏ hơn ngày PO";
                                    dataGridView1.Rows[row.Index].Cells[8].Style.BackColor = Color.Red;
                                }
                                else if (int.Parse(DELIVERY_QTY) > po_balance)
                                {
                                    dataGridView1.Rows[row.Index].Cells[8].Value = "NG - Giao hàng nhiều hơn PO";
                                    dataGridView1.Rows[row.Index].Cells[8].Style.BackColor = Color.Red;
                                }
                                else if (int.Parse(DELIVERY_QTY) <= po_balance)
                                {
                                    pro.InsertInvoice(CTR_CD, CUST_CD, EMPL_NO, G_CODE, PO_NO, DELIVERY_QTY, DELIVERY_DATE, NOCANCEL);
                                    dataGridView1.Rows[row.Index].Cells[8].Value = "OK";
                                    dataGridView1.Rows[row.Index].Cells[8].Style.BackColor = Color.Green;
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
        private void uploadInvoiceHàngLoạtToolStripMenuItem_Click(object sender, EventArgs e)
        {
            upInvoicehangloat();
        }

        public string generate_condition()
        {
            string query = "WHERE ";
            string fromdate, todate;
            fromdate = STYMD2(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day);
            todate = STYMD2(dateTimePicker2.Value.Year, dateTimePicker2.Value.Month, dateTimePicker2.Value.Day);
            string ngaythang = "PO_DATE BETWEEN '" + fromdate + "' AND '" + todate + "' ";
            if(checkBox2.Checked ==true)
            {
                ngaythang = " 1=1 ";
            }

            string code;
            if (textBox1.Text != "")
            {
                code = "AND M100.G_NAME LIKE '%" + textBox1.Text + "%' ";
            }
            else
            {
                code = "";
            }

            string empl_name = "";
            if (textBox8.Text != "")
            {
                empl_name = "AND EMPL_NAME LIKE '%" + textBox8.Text + "%' ";
            }
            else
            {
                empl_name = "";
            }

            string cust_name_kd = "";
            if (textBox9.Text != "")
            {
                cust_name_kd = "AND CUST_NAME_KD LIKE '%" + textBox9.Text + "%' ";
            }
            else
            {
                cust_name_kd = "";
            }
            string prod_type = "";

            if (textBox10.Text != "")
            {
                prod_type = "AND PROD_TYPE LIKE '%" + textBox10.Text + "%' ";
            }
            else
            {
                prod_type = "";
            }

            string po_no = "";

            if (textBox11.Text != "")
            {
                po_no = "AND ZTBPOTable.PO_NO LIKE '%" + textBox11.Text + "%' ";
            }
            else
            {
                po_no = "";
            }

            string material = "";

            if (textBox12.Text != "")
            {
                material = "AND M100.PROD_MAIN_MATERIAL LIKE '%" + textBox12.Text + "%' ";
            }
            else
            {
                material = "";
            }

            string overdue = "";

            if (textBox13.Text != "")
            {
                overdue = "AND (CASE WHEN (ZTBPOTable.RD_DATE < GETDATE()-1) AND ((ZTBPOTable.PO_QTY-AA.TotalDelivered) <>0) THEN 'OVER' ELSE 'OK' END) LIKE '%" + textBox13.Text + "%' ";
            }
            else
            {
                overdue = "";
            }
            string id = "";
            if (textBox15.Text != "")
            {
                id = "AND PO_ID=" + textBox15.Text;
            }
            else
            {
                id = "";
            }

            string chiton = "";
            if (checkBox3.Checked == true)
            {
                chiton = "AND (ZTBPOTable.PO_QTY - AA.TotalDelivered) <>0";
            }
            else
            {
                chiton = "";
            }

            string cmscode = "";
            if (textBox16.Text != "")
            {
                cmscode = "AND M100.G_CODE='" + textBox16.Text + "'";
            }
            else
            {
                cmscode = "";
            }

            query += ngaythang + code + empl_name + cust_name_kd + prod_type + po_no + material + overdue+ id + chiton + cmscode;
            return query;

        }


        public string generate_condition_chokiem()
        {
            string query = "WHERE ";
            string fromdate, todate;
            fromdate = STYMD2(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day);
            todate = STYMD2(dateTimePicker2.Value.Year, dateTimePicker2.Value.Month, dateTimePicker2.Value.Day);
            //MessageBox.Show(fromdate);
            string ngaythang = "UPDATE_DATE = '" + fromdate + "'";
            
            string code;
            if (textBox1.Text != "")
            {
                code = "AND M100.G_NAME LIKE '%" + textBox1.Text + "%' ";
            }
            else
            {
                code = "";
            }

      
            string id = "";
            if (textBox15.Text != "")
            {
                id = "AND WI_ID=" + textBox15.Text;
            }
            else
            {
                id = "";
            }

            string cmscode = "";
            if (textBox16.Text != "")
            {
                cmscode = "AND M100.G_CODE='" + textBox16.Text + "'";
            }
            else
            {
                cmscode = "";
            }
            string calamviec = " AND ZTB_WAIT_INSPECT.CALAMVIEC='DEM'";
            
            query += ngaythang + code  + id + cmscode + calamviec;
            return query;
        }

        public string generate_condition_BTP()
        {
            string query = "WHERE ";
            string fromdate, todate;
            fromdate = STYMD2(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day);
            todate = STYMD2(dateTimePicker2.Value.Year, dateTimePicker2.Value.Month, dateTimePicker2.Value.Day);
            //MessageBox.Show(fromdate);
            string ngaythang = "UPDATE_DATE = '" + fromdate + "'";

            string code;
            if (textBox1.Text != "")
            {
                code = "AND M100.G_NAME LIKE '%" + textBox1.Text + "%' ";
            }
            else
            {
                code = "";
            }


            string id = "";
            if (textBox15.Text != "")
            {
                id = "AND HG_ID=" + textBox15.Text;
            }
            else
            {
                id = "";
            }

            string cmscode = "";
            if (textBox16.Text != "")
            {
                cmscode = "AND M100.G_CODE='" + textBox16.Text + "'";
            }
            else
            {
                cmscode = "";
            }
            

            query += ngaythang + code + id + cmscode;
            return query;
        }





        public string generate_condition_invoice()
        {
            string query = "WHERE ";
            string fromdate, todate;
            fromdate = STYMD2(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day);
            todate = STYMD2(dateTimePicker2.Value.Year, dateTimePicker2.Value.Month, dateTimePicker2.Value.Day);
            //MessageBox.Show(fromdate);
            string ngaythang = "DELIVERY_DATE BETWEEN '" + fromdate + "' AND '" + todate + "' ";
            if (checkBox2.Checked == true)
            {
                ngaythang = "1=1 ";
            }

            string code;
            if (textBox1.Text != "")
            {
                code = "AND M100.G_NAME LIKE '%" + textBox1.Text + "%' ";
            }
            else
            {
                code = "";
            }

            string empl_name = "";
            if (textBox8.Text != "")
            {
                empl_name = "AND M010.EMPL_NAME LIKE '%" + textBox8.Text + "%' ";
            }
            else
            {
                empl_name = "";
            }

            string cust_name_kd = "";
            if (textBox9.Text != "")
            {
                cust_name_kd = "AND M110.CUST_NAME_KD LIKE '%" + textBox9.Text + "%' ";
            }
            else
            {
                cust_name_kd = "";
            }
            string prod_type = "";

            if (textBox10.Text != "")
            {
                prod_type = "AND M100.PROD_TYPE LIKE '%" + textBox10.Text + "%' ";
            }
            else
            {
                prod_type = "";
            }

            string po_no = "";

            if (textBox11.Text != "")
            {
                po_no = "AND ZTBPOTable.PO_NO LIKE '%" + textBox11.Text + "%' ";
            }
            else
            {
                po_no = "";
            }

            string material = "";

            if (textBox12.Text != "")
            {
                material = "AND M100.PROD_MAIN_MATERIAL LIKE '%" + textBox12.Text + "%' ";
            }
            else
            {
                material = "";
            }

            string invoice_no = "";

            if (textBox14.Text != "")
            {
                invoice_no = "AND ZTBDelivery.INVOICE_NO LIKE '%" + textBox14.Text + "%' ";
            }
            else
            {
                invoice_no = "";
            }

            string id = "";
            if (textBox15.Text != "")
            {
                id = "AND DELIVERY_ID=" + textBox15.Text;
            }
            else
            {
                id = "";
            }

            string cmscode = "";
            if (textBox16.Text != "")
            {
                cmscode = "AND M100.G_CODE='" + textBox16.Text + "'";
            }
            else
            {
                cmscode = "";
            }

            query += ngaythang + code + empl_name + cust_name_kd + prod_type + po_no + material + invoice_no + id + cmscode;
            return query;
        }
        public int GetWeekNumber(DateTime datetime )
        {
            CultureInfo ciCurr = CultureInfo.CurrentCulture;
            int weekNum = ciCurr.Calendar.GetWeekOfYear(datetime, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Sunday);
            return weekNum;
        }

        public string generate_condition_fsct()
        {
            string query = "WHERE ";
            string fromdate, todate;
            fromdate = STYMD2(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day);
            todate = STYMD2(dateTimePicker2.Value.Year, dateTimePicker2.Value.Month, dateTimePicker2.Value.Day);
            int tuan1 = GetWeekNumber(dateTimePicker1.Value);
            int tuan2 = GetWeekNumber(dateTimePicker2.Value);
            string tuan = "FCSTWEEKNO BETWEEN " + tuan1 +  " AND  " + tuan2;
           

            if (checkBox2.Checked == true)
            {
                tuan = "1=1 ";
            }

            string code;
            if (textBox1.Text != "")
            {
                code = "AND M100.G_NAME LIKE '%" + textBox1.Text + "%' ";
            }
            else
            {
                code = "";
            }

            string empl_name = "";
            if (textBox8.Text != "")
            {
                empl_name = "AND M010.EMPL_NAME LIKE '%" + textBox8.Text + "%' ";
            }
            else
            {
                empl_name = "";
            }

            string cust_name_kd = "";
            if (textBox9.Text != "")
            {
                cust_name_kd = "AND M110.CUST_NAME_KD LIKE '%" + textBox9.Text + "%' ";
            }
            else
            {
                cust_name_kd = "";
            }
            string prod_type = "";

            if (textBox10.Text != "")
            {
                prod_type = "AND M100.PROD_TYPE LIKE '%" + textBox10.Text + "%' ";
            }
            else
            {
                prod_type = "";
            }

            string year = "";     
            
            year = "AND ZTBFCSTTB.FCSTYEAR= "+ dateTimePicker1.Value.Year.ToString() + "";
            if (checkBox2.Checked == true)
            {
                year = "AND 1=1 ";
            }

            string material = "";

            if (textBox12.Text != "")
            {
                material = "AND M100.PROD_MAIN_MATERIAL LIKE '%" + textBox12.Text + "%' ";
            }
            else
            {
                material = "";
            }

            string id = "";
            if (textBox15.Text != "")
            {
                id = "AND FCST_ID=" + textBox15.Text;
            }
            else
            {
                id = "";
            }
            string cmscode = "";
            if (textBox16.Text !="")
            {
                cmscode = "AND M100.G_CODE='" + textBox16.Text + "'";
            }
            else
            {
                cmscode = "";
            }

            query += tuan + code + empl_name + cust_name_kd + prod_type + year + material+id + cmscode;
            return query;


        }


        public string generate_condition_plan()
        {
            string query = "WHERE ";
            string fromdate, todate;
            fromdate = STYMD2(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day);
            todate = STYMD2(dateTimePicker2.Value.Year, dateTimePicker2.Value.Month, dateTimePicker2.Value.Day);            
            string tuan = "ZTBPLANTB.PLAN_DATE BETWEEN '" + fromdate + "' AND  '" + todate + "'";
           // MessageBox.Show(fromdate);
            if (checkBox2.Checked == true)
            {
                tuan = "1=1 ";
            }
            string code;
            if (textBox1.Text != "")
            {
                code = "AND M100.G_NAME LIKE '%" + textBox1.Text + "%' ";
            }
            else
            {
                code = "";
            }

            string empl_name = "";
            if (textBox8.Text != "")
            {
                empl_name = "AND M010.EMPL_NAME LIKE '%" + textBox8.Text + "%' ";
            }
            else
            {
                empl_name = "";
            }

            string cust_name_kd = "";
            if (textBox9.Text != "")
            {
                cust_name_kd = "AND M110.CUST_NAME_KD LIKE '%" + textBox9.Text + "%' ";
            }
            else
            {
                cust_name_kd = "";
            }
            string prod_type = "";

            if (textBox10.Text != "")
            {
                prod_type = "AND M100.PROD_TYPE LIKE '%" + textBox10.Text + "%' ";
            }
            else
            {
                prod_type = "";
            }

            
            string material = "";

            if (textBox12.Text != "")
            {
                material = "AND M100.PROD_MAIN_MATERIAL LIKE '%" + textBox12.Text + "%' ";
            }
            else
            {
                material = "";
            }

            string id = "";
            if (textBox15.Text != "")
            {
                id = "AND PLAN_ID=" + textBox15.Text;
            }
            else
            {
                id = "";
            }
            string cmscode = "";
            if (textBox16.Text != "")
            {
                cmscode = "AND M100.G_CODE='" + textBox16.Text + "'";
            }
            else
            {
                cmscode = "";
            }

            query += tuan + code + empl_name + cust_name_kd + prod_type +  material+id+cmscode;
            return query;

        }


        private void traPOToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                button4.Enabled = true;
                string fromdate, todate;
                fromdate = STYMD(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day);
                todate = STYMD(dateTimePicker2.Value.Year, dateTimePicker2.Value.Month, dateTimePicker2.Value.Day);
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();
                DataTable dttong = new DataTable();
                //MessageBox.Show(generate_condition());
                dt = pro.traPO(generate_condition());
                dttong = pro.traPOTotal(generate_condition());

                if (dt.Rows.Count > 0)
                {
                    dataGridView1.DataSource = dt;
                    button3.Enabled = true;
                    //button5.Enabled = true; 


                    textBox2.Text = dttong.Rows[0]["PO_QTY"].ToString();
                    textBox2.Text = string.Format("{0:#,##0 EA}", double.Parse(textBox2.Text));
                    textBox4.Text = dttong.Rows[0]["TOTAL_DELIVERED"].ToString();
                    textBox4.Text = string.Format("{0:#,##0 EA}", double.Parse(textBox4.Text));
                    textBox6.Text = dttong.Rows[0]["PO_BALANCE"].ToString();
                    textBox6.Text = string.Format("{0:#,##0 EA}", double.Parse(textBox6.Text));

                    textBox3.Text = dttong.Rows[0]["PO_AMOUNT"].ToString();
                    textBox3.Text = string.Format("{0:#,##0.00 $}", double.Parse(textBox3.Text));
                    textBox5.Text = dttong.Rows[0]["DELIVERED_AMOUNT"].ToString();
                    textBox5.Text = string.Format("{0:#,##0.00 $}", double.Parse(textBox5.Text));
                    textBox7.Text = dttong.Rows[0]["BALANCE_AMOUNT"].ToString();
                    textBox7.Text = string.Format("{0:#,##0.00 $}", double.Parse(textBox7.Text));

                    textBox2.Font = new Font("Microsoft Sans Serif", 15.0f);
                    textBox4.Font = new Font("Microsoft Sans Serif", 15.0f);
                    textBox6.Font = new Font("Microsoft Sans Serif", 15.0f);
                    textBox3.Font = new Font("Microsoft Sans Serif", 15.0f);
                    textBox5.Font = new Font("Microsoft Sans Serif", 15.0f);
                    textBox7.Font = new Font("Microsoft Sans Serif", 15.0f);


                    MessageBox.Show("Đã load  " + dt.Rows.Count + "dòng");
                }
                else
                {
                    MessageBox.Show("Không có kết quả nào");
                }
                button8.Enabled = true;

            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi kết nối ! " + ex.ToString());
            }
        }

       

        public void traInvoice()
        {
            try
            {
                bom_flag = 0;
                fcst_flag = 0;
                invoice_flag = 1;
                khgh_flag = 0;
                po_flag = 0;
                ycsx_flag = 0;

                button4.Enabled = true;
                string fromdate, todate;
                fromdate = STYMD(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day);
                todate = STYMD(dateTimePicker2.Value.Year, dateTimePicker2.Value.Month, dateTimePicker2.Value.Day);
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();
                DataTable dttong = new DataTable();
                dt = pro.traInvoice(generate_condition_invoice());
                dttong = pro.traInvoiceTotal(generate_condition_invoice());


                if (dt.Rows.Count > 0)
                {
                    this.dataGridView1.DataSource = null;
                    this.dataGridView1.Columns.Clear();
                    dataGridView1.DataSource = dt;
                    //DoubleBuffered(dataGridView1,true);
                    

                    setRowNumber(dataGridView1);
                    button3.Enabled = true;
                    //button5.Enabled = true; 


                    textBox2.Text = "0";
                    textBox2.Text = string.Format("{0:#,##0 EA}", double.Parse(textBox2.Text));
                    textBox4.Text = dttong.Rows[0]["DELIVERED_QTY"].ToString();
                    textBox4.Text = string.Format("{0:#,##0 EA}", double.Parse(textBox4.Text));
                    textBox6.Text = "0";
                    textBox6.Text = string.Format("{0:#,##0 EA}", double.Parse(textBox6.Text));

                    textBox3.Text = "0";
                    textBox3.Text = string.Format("{0:#,##0.00 $}", double.Parse(textBox3.Text));
                    textBox5.Text = dttong.Rows[0]["DELIVERED_AMOUNT"].ToString();
                    textBox5.Text = string.Format("{0:#,##0.00 $}", double.Parse(textBox5.Text));
                    textBox7.Text = "0";
                    textBox7.Text = string.Format("{0:#,##0.00 $}", double.Parse(textBox7.Text));

                    textBox2.Font = new Font("Microsoft Sans Serif", 15.0f);
                    textBox4.Font = new Font("Microsoft Sans Serif", 15.0f);
                    textBox6.Font = new Font("Microsoft Sans Serif", 15.0f);
                    textBox3.Font = new Font("Microsoft Sans Serif", 15.0f);
                    textBox5.Font = new Font("Microsoft Sans Serif", 15.0f);
                    textBox7.Font = new Font("Microsoft Sans Serif", 15.0f);

                    formatDataGridViewtraInvoice1(dataGridView1);
                    MessageBox.Show("Đã load  " + dt.Rows.Count + "dòng");
                    po_flag = 0;
                    invoice_flag = 1;
                    khgh_flag = 0;
                    fcst_flag = 0;
                    ycsx_flag = 0;
                }
                else
                {
                    MessageBox.Show("Không có kết quả nào");
                }
                button8.Enabled = true;

            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi kết nối ! " + ex.ToString());
            }
        }

        public void traInvoice_Async(DoWorkEventArgs e)
        {
            try
            {               
                button4.Enabled = true;
                string fromdate, todate;
                fromdate = STYMD(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day);
                todate = STYMD(dateTimePicker2.Value.Year, dateTimePicker2.Value.Month, dateTimePicker2.Value.Day);
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();
                DataTable dttong = new DataTable();
                dtgv1_data = pro.traInvoice(generate_condition_invoice());
                dtgv1_total_data = pro.traInvoiceTotal(generate_condition_invoice());
                button8.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi kết nối ! " + ex.ToString());
            }
        }

        private void button24_Click(object sender, EventArgs e)
        {
            dataGridView1.Show();
            traInvoice();
        }

        private void taoChiThiSXToolStripMenuItem_Click(object sender, EventArgs e)
        {
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

                        //dataGridView1.Rows[rowindex].Cells[columnindex].Value.ToString();

                        var selectedRows = dataGridView1.SelectedRows
                       .OfType<DataGridViewRow>()
                       .Where(row => !row.IsNewRow)
                       .ToArray();


                        progressBar1.Minimum = 0; //Đặt giá trị nhỏ nhất cho ProgressBar
                        progressBar1.Maximum = dataGridView1.SelectedRows.Count; //Đặt giá trị lớn nhất cho ProgressBar
                        int startprogress = 0;

                        foreach (var row in selectedRows)
                        {
                            string ycsxno = row.Cells[3].Value.ToString();
                            dt = pro.getFullInfo(ycsxno);
                            if (file != "")
                            {
                                ExcelFactory.editFileExcelQLSX(file, dt, checkBox1, saveycsxpath);
                                startprogress = startprogress + 1;
                                label4.Text = "Progress: " + startprogress + "/" + dataGridView1.SelectedRows.Count;
                                progressBar1.Value = startprogress;
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

        private void nhậpExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            po_flag = 0;
            invoice_flag = 0;
            khgh_flag = 0;
            fcst_flag = 0;
            ycsx_flag = 0;
            import_excel_flag = 0;
            check_po_flag = 0;
            up_po_flag = 0;
            check_invoice_flag = 0;
            up_invoice_flag = 0;
            check_fcst_flag = 0;
            up_fcst_flag = 0;
            check_plan_flag = 0;
            up_plan_flag = 0;

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
                    button4.Enabled = true;
                    button3.Enabled = true;
                    button16.Enabled = true;
                    import_excel_flag = 0;                  

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void xuấtExcelCủaBảngHiệnTạiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridView1 == null || dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("Không có dữ liệu!");
            }
            else
            {
                ExcelFactory.writeToExcelFile(dataGridView1);
            }
        }

      

        private void uploadKHGHHàngLoạtToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //insertKHGH(1);

        }

        private void tạo1POMơiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            POForm poform = new POForm();
            poform.loginIDpoForm = LoginID;
            poform.Show();
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
                            string G_CODE, G_CODE_KD, PROD_TYPE, PROD_MODEL, PROD_PROJECT, PROD_SIZE, PROD_MAIN_MATERIAL;

                            G_CODE = Convert.ToString(row.Cells[0].Value);
                            G_CODE_KD = Convert.ToString(row.Cells[1].Value);
                            PROD_TYPE = Convert.ToString(row.Cells[2].Value);
                            PROD_MODEL = Convert.ToString(row.Cells[3].Value);
                            PROD_PROJECT = Convert.ToString(row.Cells[4].Value);
                            PROD_SIZE = Convert.ToString(row.Cells[5].Value);
                            PROD_MAIN_MATERIAL = Convert.ToString(row.Cells[6].Value);

                            ProductBLL pro = new ProductBLL();
                            DataTable dt = new DataTable();
                            pro.updateInfo(G_CODE, G_CODE_KD, PROD_TYPE, PROD_MODEL, PROD_PROJECT, PROD_MAIN_MATERIAL);

                        }
                    }

                    MessageBox.Show("Đã hoàn thành update thông tin code hàng loạt");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }

        }

        private void upThôngTinHàngLoạtToolStripMenuItem_Click(object sender, EventArgs e)
        {
            updatecodeinfo();
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
                            CUST_NAME_KD = Convert.ToString(row.Cells[1].Value);
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
        private void upThôngTinHàngLoạtToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            uploadcustomerinfor();
        }

        public void traFCST()
        {
            po_flag = 0;
            invoice_flag = 0;
            khgh_flag = 0;
            fcst_flag = 1;
            ycsx_flag = 0;
            try
            {
                
                //MessageBox.Show(""+GetWeekNumber(dateTimePicker1.Value));
                button4.Enabled = true;
                string fromdate, todate;
                fromdate = STYMD(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day);
                todate = STYMD(dateTimePicker2.Value.Year, dateTimePicker2.Value.Month, dateTimePicker2.Value.Day);
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();
                DataTable dttong = new DataTable();
                //MessageBox.Show(generate_condition());
                dt = pro.traFCST(generate_condition_fsct());
                //dttong = pro.traPOTotal(generate_condition());

                if (dt.Rows.Count > 0)
                {
                    this.dataGridView1.DataSource = null;
                    this.dataGridView1.Columns.Clear();
                    dataGridView1.DataSource = dt;
                    setRowNumber(dataGridView1);
                    formatDataGridViewtraFCST1(dataGridView1);
                    MessageBox.Show("Đã load  " + dt.Rows.Count + "dòng");
                }
                else
                {
                    MessageBox.Show("Không có kết quả nào");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi kết nối ! " + ex.ToString());
            }
        }

        public void traFCST_Async(DoWorkEventArgs e)
        {          
            try
            {
                button4.Enabled = true;
                string fromdate, todate;
                fromdate = STYMD(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day);
                todate = STYMD(dateTimePicker2.Value.Year, dateTimePicker2.Value.Month, dateTimePicker2.Value.Day);
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();
                DataTable dttong = new DataTable();
               
                dtgv1_data = pro.traFCST(generate_condition_fsct());
             
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi kết nối ! " + ex.ToString());
            }
        }

        private void button23_Click(object sender, EventArgs e)
        {
            dataGridView1.Show();
            traFCST();
        }

        public void traPLAN()
        {
            try
            {
                bom_flag = 0;
                fcst_flag = 0;
                invoice_flag = 0;
                khgh_flag = 1;
                po_flag = 0;
                ycsx_flag = 0;
                //MessageBox.Show(""+GetWeekNumber(dateTimePicker1.Value));
                button4.Enabled = true;
                string fromdate, todate;
                fromdate = STYMD(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day);
                todate = STYMD(dateTimePicker2.Value.Year, dateTimePicker2.Value.Month, dateTimePicker2.Value.Day);
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();
                DataTable dttong = new DataTable();
                //MessageBox.Show(generate_condition());
                dt = pro.traPlan(generate_condition_plan());
                //dttong = pro.traPOTotal(generate_condition());

                if (dt.Rows.Count > 0)
                {
                    this.dataGridView1.DataSource = null;
                    this.dataGridView1.Columns.Clear();
                    dataGridView1.DataSource = dt;
                    setRowNumber(dataGridView1);
                    formatDataGridViewtraPlan1(dataGridView1);
                    MessageBox.Show("Đã load  " + dt.Rows.Count + "dòng");
                }
                else
                {
                    MessageBox.Show("Không có kết quả nào");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Không có dữ liệu");
            }
        }

        public void traPLAN_Async(DoWorkEventArgs e)
        {
            try
            {               
                button4.Enabled = true;
                string fromdate, todate;
                fromdate = STYMD(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day);
                todate = STYMD(dateTimePicker2.Value.Year, dateTimePicker2.Value.Month, dateTimePicker2.Value.Day);
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();
                DataTable dttong = new DataTable();
                //MessageBox.Show(generate_condition());
                dtgv1_data = pro.traPlan(generate_condition_plan());
                //dttong = pro.traPOTotal(generate_condition());
            }
            catch (Exception ex)
            {
                MessageBox.Show("Không có dữ liệu");
            }
        }

        private void button22_Click(object sender, EventArgs e)
        {
            dataGridView1.Show();
            traPLAN();
        }

        private void dataGridView1_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
           

        }

        public void xoaInvoice()
        {
            if(invoice_flag==1)
            {
                if (MessageBox.Show("Bạn thực sự muốn xóa Invoice?", "Xóa Invoice ?", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    try
                    {
                        ProductBLL pro = new ProductBLL();
                        DataTable dt = new DataTable();

                        var selectedRows = dataGridView1.SelectedRows
                       .OfType<DataGridViewRow>()
                       .Where(row => !row.IsNewRow)
                       .ToArray();

                        progressBar1.Minimum = 0; //Đặt giá trị nhỏ nhất cho ProgressBar
                        progressBar1.Maximum = dataGridView1.SelectedRows.Count; //Đặt giá trị lớn nhất cho ProgressBar
                        int startprogress = 0;

                        foreach (var row in selectedRows)
                        {
                            string delivery_id = row.Cells["DELIVERY_ID"].Value.ToString(); // cell 15
                            dt = pro.DeleteInvoice(delivery_id);
                            pro.writeHistory("002", LoginID, "DELIVERY TABLE", "XOA", "XOA INVOICE", "" + delivery_id);

                            startprogress = startprogress + 1;
                            label4.Text = "Progress: " + startprogress + "/" + dataGridView1.SelectedRows.Count;
                            progressBar1.Value = startprogress;
                        }

                        progressBar1.Value = 0;
                        dataGridView1.ClearSelection();
                        MessageBox.Show("Xóa Invoices thành công !");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }


                }
            }
            else
            {
                MessageBox.Show("Không phải bảng invoice nên không xóa được");
            }
            
                
            
        }
        private void xóaINVOICEĐãChọnToolStripMenuItem_Click(object sender, EventArgs e)
        {
            xoaInvoice();
        }


        public void checkInvoice()
        {
            int total_flag = po_flag + invoice_flag + fcst_flag + khgh_flag;
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
                            ProductBLL pro = new ProductBLL();
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

                                DateTime dlidate = DateTime.Parse(DELIVERY_DATE);          // convert delivery date to datetime                     
                                int check_date = checkDate(dlidate); // so sanh delivery date vs today
                                DateTime podate = DateTime.Parse(pro.getPODDate(G_CODE, PO_NO)); // get podate and convert to datetime
                                int checkinvoicedatevspodate = checkInvoicevsPODate(dlidate, podate);
                               


                                if ((CUST_CD == "") || (EMPL_NO == "") || (G_CODE == "") || (PO_NO == "") || (DELIVERY_QTY == "") || (DELIVERY_DATE == "") || (NOCANCEL == ""))
                                {
                                    MessageBox.Show("Thông tin của sản phẩm: " + G_CODE + "có thông tin trống, sẽ ko thêm PO này !");
                                }
                                else
                                {
                                    
                                    int po_balance = pro.checkPOBalance(CUST_CD, G_CODE, PO_NO);
                                    
                                    if(check_date==0)
                                    {
                                        dataGridView1.Rows[row.Index].Cells[8].Value = "NG - Ngày Invoice lớn hơn ngày hiện tại";
                                        dataGridView1.Rows[row.Index].Cells[8].Style.BackColor = Color.Red;
                                    }
                                    else if(checkinvoicedatevspodate==0)
                                    {
                                        dataGridView1.Rows[row.Index].Cells[8].Value = "NG - Ngày Invoice Nhỏ hơn ngày PO";
                                        dataGridView1.Rows[row.Index].Cells[8].Style.BackColor = Color.Red;
                                    }
                                    else if(int.Parse(DELIVERY_QTY) > po_balance)
                                    {
                                        dataGridView1.Rows[row.Index].Cells[8].Value = "NG - Giao hàng nhiều hơn PO";
                                        dataGridView1.Rows[row.Index].Cells[8].Style.BackColor = Color.Red;
                                    }
                                    else if (int.Parse(DELIVERY_QTY) <= po_balance)
                                    {
                                        //pro.InsertInvoice(CTR_CD, CUST_CD, EMPL_NO, G_CODE, PO_NO, DELIVERY_QTY, DELIVERY_DATE, NOCANCEL);
                                        dataGridView1.Rows[row.Index].Cells[8].Value = "OK";
                                        dataGridView1.Rows[row.Index].Cells[8].Style.BackColor = Color.LightGreen;
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

        public void checkInvoice_Async(DoWorkEventArgs e)
        {           
            if (import_excel_flag == 1)
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
                            ProductBLL pro = new ProductBLL();
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

                                DateTime dlidate = DateTime.Parse(DELIVERY_DATE);          // convert delivery date to datetime                     
                                int check_date = checkDate(dlidate); // so sanh delivery date vs today
                                DateTime podate = DateTime.Parse(pro.getPODDate(G_CODE, PO_NO)); // get podate and convert to datetime
                                int checkinvoicedatevspodate = checkInvoicevsPODate(dlidate, podate);



                                if ((CUST_CD == "") || (EMPL_NO == "") || (G_CODE == "") || (PO_NO == "") || (DELIVERY_QTY == "") || (DELIVERY_DATE == "") || (NOCANCEL == ""))
                                {
                                    MessageBox.Show("Thông tin của sản phẩm: " + G_CODE + "có thông tin trống, sẽ ko thêm PO này !");
                                }
                                else
                                {

                                    int po_balance = pro.checkPOBalance(CUST_CD, G_CODE, PO_NO);

                                    if (check_date == 0)
                                    {
                                        dataGridView1.Rows[row.Index].Cells[8].Value = "NG - Ngày Invoice lớn hơn ngày hiện tại";
                                        dataGridView1.Rows[row.Index].Cells[8].Style.BackColor = Color.Red;
                                    }
                                    else if (checkinvoicedatevspodate == 0)
                                    {
                                        dataGridView1.Rows[row.Index].Cells[8].Value = "NG - Ngày Invoice Nhỏ hơn ngày PO";
                                        dataGridView1.Rows[row.Index].Cells[8].Style.BackColor = Color.Red;
                                    }
                                    else if (int.Parse(DELIVERY_QTY) > po_balance)
                                    {
                                        dataGridView1.Rows[row.Index].Cells[8].Value = "NG - Giao hàng nhiều hơn PO";
                                        dataGridView1.Rows[row.Index].Cells[8].Style.BackColor = Color.Red;
                                    }
                                    else if (int.Parse(DELIVERY_QTY) <= po_balance)
                                    {
                                        //pro.InsertInvoice(CTR_CD, CUST_CD, EMPL_NO, G_CODE, PO_NO, DELIVERY_QTY, DELIVERY_DATE, NOCANCEL);
                                        dataGridView1.Rows[row.Index].Cells[8].Value = "OK";
                                        dataGridView1.Rows[row.Index].Cells[8].Style.BackColor = Color.LightGreen;
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

        private void checkInvoiceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            checkInvoice();                
        }

        public void checkPO()
        {
            int total_flag = po_flag + invoice_flag + fcst_flag + khgh_flag;
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

                                DateTime podate = DateTime.Parse(PO_DATE);
                                int check_date = checkDate(podate);

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
                                    else if(check_date == 0)
                                    {
                                        dataGridView1.Rows[row.Index].Cells[8].Value = "NG - Ngày PO lớn hơn ngày hiện tại";
                                        dataGridView1.Rows[row.Index].Cells[8].Style.BackColor = Color.Red;
                                    }
                                    else
                                    {
                                        dataGridView1.Rows[row.Index].Cells[8].Value = "NG - Đã tồn tại PO";
                                        dataGridView1.Rows[row.Index].Cells[8].Style.BackColor = Color.Red;
                                    }

                                }

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

        public void checkPO_Async(DoWorkEventArgs e)
        {            
            if (import_excel_flag == 1)
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

                                DateTime podate = DateTime.Parse(PO_DATE);
                                int check_date = checkDate(podate);
                                
                                if ((CUST_CD == "") || (EMPL_NO == "") || (G_CODE == "") || (PO_NO == "") || (PO_QTY == "") || (PO_DATE == "") || (RD_DATE == "") || (PROD_PRICE == ""))
                                {
                                    MessageBox.Show("Thông tin của sản phẩm: " + G_CODE + "có thông tin trống, sẽ ko thêm PO này !");
                                }
                                else
                                {
                                    ProductBLL pro = new ProductBLL();
                                    string checkUSEYN = pro.checkM100UseYN(G_CODE);

                                    if (pro.checkPOExist(CUST_CD, G_CODE, PO_NO) != -1)
                                    {
                                        dataGridView1.Rows[row.Index].Cells[8].Value = "NG - Đã tồn tại PO";
                                        dataGridView1.Rows[row.Index].Cells[8].Style.BackColor = Color.Red;                                        
                                    }
                                    else if (check_date == 0)
                                    {
                                        dataGridView1.Rows[row.Index].Cells[8].Value = "NG - Ngày PO lớn hơn ngày hiện tại";
                                        dataGridView1.Rows[row.Index].Cells[8].Style.BackColor = Color.Red;
                                    }
                                    else if(checkUSEYN=="N")
                                    {
                                        dataGridView1.Rows[row.Index].Cells[8].Value = "NG - Ver này đã bị khoá, check lại ver final";
                                        dataGridView1.Rows[row.Index].Cells[8].Style.BackColor = Color.Red;
                                    }
                                    else
                                    {
                                        dataGridView1.Rows[row.Index].Cells[8].Value = "OK";
                                        dataGridView1.Rows[row.Index].Cells[8].Style.BackColor = Color.LightGreen;
                                    }

                                }

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


        private void checkPOToolStripMenuItem_Click(object sender, EventArgs e)
        {
            checkPO();
            
        }



        public void insertKHGH(int check, DoWorkEventArgs e)
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

                                if (D1.IndexOf("-", 0) == 0 || D2.IndexOf("-", 0) == 0 || D3.IndexOf("-", 0) == 0 || D4.IndexOf("-", 0) == 0 || D5.IndexOf("-", 0) == 0 || D6.IndexOf("-", 0) == 0 || D7.IndexOf("-", 0) == 0 || D8.IndexOf("-", 0) == 0 )
                                {
                                    MessageBox.Show("Không được phép có giá trị âm, bỏ qua dòng này");
                                    dataGridView1.Rows[row.Index].Cells[29].Value = "NG - Giá trị plan âm";
                                    dataGridView1.Rows[row.Index].Cells[29].Style.BackColor = Color.Red;
                                }
                                else if ((EMPL_NO == "") || (CUST_CD == "") || (G_CODE == "") || (PLAN_DATE == "")  || (D1 == "") || (D2 == "") || (D3 == "") || (D4 == "") || (D5 == "") || (D6 == "") || (D7 == "") || (D8 == ""))
                                {
                                    MessageBox.Show("Thông tin của sản phẩm: " + G_CODE + "có thông tin trống, sẽ ko thêm KHGH này !");
                                }
                                else
                                {
                                    ProductBLL pro = new ProductBLL();
                                    string checkUSEYN = pro.checkM100UseYN(G_CODE);
                                    if (pro.checkKHGHExist(CUST_CD, G_CODE, PLAN_DATE) != -1)
                                    {
                                        dataGridView1.Rows[row.Index].Cells[13].Value = "NG - Đã tồn tại KHGH";
                                        dataGridView1.Rows[row.Index].Cells[13].Style.BackColor = Color.Red;
                                    }
                                    else if(checkUSEYN == "N")
                                    {
                                        dataGridView1.Rows[row.Index].Cells[13].Value = "NG - Ver này đã bị khoá, check lại ver code này";
                                        dataGridView1.Rows[row.Index].Cells[13].Style.BackColor = Color.Red;
                                    }
                                    else
                                    {
                                        pro.InsertPlan(CTR_CD, EMPL_NO, CUST_CD, G_CODE, PLAN_DATE, D1, D2, D3, D4, D5, D6, D7, D8, REMARK);
                                        pro.writeHistory(CTR_CD, LoginID, "PLAN TABLE", "THEM", "THEM PLAN GIAO HANG", "0");
                                        dataGridView1.Rows[row.Index].Cells[13].Value = "OK";
                                        dataGridView1.Rows[row.Index].Cells[13].Style.BackColor = Color.LightGreen;
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
                                    string checkUSEYN = pro.checkM100UseYN(G_CODE);
                                    if (pro.checkKHGHExist(CUST_CD, G_CODE, PLAN_DATE) != -1)
                                    {
                                        dataGridView1.Rows[row.Index].Cells[13].Value = "NG - Đã tồn tại KHGH";
                                        dataGridView1.Rows[row.Index].Cells[13].Style.BackColor = Color.Red;
                                    }
                                    else if (checkUSEYN == "N")
                                    {
                                        dataGridView1.Rows[row.Index].Cells[13].Value = "NG - Ver này đã bị khoá, check lại ver code này";
                                        dataGridView1.Rows[row.Index].Cells[13].Style.BackColor = Color.Red;
                                    }
                                    else
                                    {                                        
                                        dataGridView1.Rows[row.Index].Cells[13].Value = "OK";
                                        dataGridView1.Rows[row.Index].Cells[13].Style.BackColor = Color.LightGreen;
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



        public void insertFCST(int check, DoWorkEventArgs e)
        {
            if(check == 1)
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
                                string CUST_CD, EMPL_NO, G_CODE,  YEAR, WEEKNO, PROD_PRICE, W1, W2, W3, W4, W5, W6, W7, W8, W9, W10, W11, W12, W13, W14, W15, W16, W17, W18, W19, W20, W21, W22;
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
                                else  if ((EMPL_NO == "") || (CUST_CD == "") || (G_CODE == "") || (PROD_PRICE == "") || (YEAR == "") || (WEEKNO == "") || (W1 == "") || (W2 == "") || (W3 == "") || (W4 == "") || (W5 == "") || (W6 == "") || (W7 == "") || (W8 == "") || (W9 == "") || (W10 == "") || (W11 == "") || (W12 == "") || (W13 == "") || (W14 == "") || (W15 == "") || (W16 == "") || (W17 == "") || (W18 == "") || (W19 == "") || (W20 == "") || (W21 == "") || (W22 == ""))
                                {
                                    MessageBox.Show("Thông tin của sản phẩm: " + G_CODE + "có thông tin trống, sẽ ko thêm PO này !");
                                }
                                else
                                {
                                    ProductBLL pro = new ProductBLL();
                                    string checkUSEYN = pro.checkM100UseYN(G_CODE);
                                    if (pro.checkFCSTExist(CUST_CD, G_CODE, YEAR,WEEKNO) != -1)
                                    {
                                        dataGridView1.Rows[row.Index].Cells[29].Value = "NG - Đã tồn tại FCST";
                                        dataGridView1.Rows[row.Index].Cells[29].Style.BackColor = Color.Red;
                                    }
                                    else if(checkUSEYN == "N")
                                    {
                                        dataGridView1.Rows[row.Index].Cells[29].Value = "NG - Ver này đã bị khoá, check lại ver final đi";
                                        dataGridView1.Rows[row.Index].Cells[29].Style.BackColor = Color.Red;
                                    }
                                    else
                                    {
                                        pro.InsertFCST(CTR_CD, EMPL_NO, CUST_CD, G_CODE, PROD_PRICE, YEAR, WEEKNO, W1, W2, W3, W4, W5, W6, W7, W8, W9, W10, W11, W12, W13, W14, W15, W16, W17, W18, W19, W20, W21, W22);
                                        dataGridView1.Rows[row.Index].Cells[29].Value = "OK";
                                        dataGridView1.Rows[row.Index].Cells[29].Style.BackColor = Color.LightGreen;
                                    }


                                }

                            }
                        }

                        MessageBox.Show("Đã hoàn thành thêm FCST hàng loạt");
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
                                else if ((EMPL_NO == "") || (CUST_CD == "") || (G_CODE == "") || (PROD_PRICE == "") || (YEAR == "") || (WEEKNO == "") || (W1 == "") || (W2 == "") || (W3 == "") || (W4 == "") || (W5 == "") || (W6 == "") || (W7 == "") || (W8 == "") || (W9 == "") || (W10 == "") || (W11 == "") || (W12 == "") || (W13 == "") || (W14 == "") || (W15 == "") || (W16 == "") || (W17 == "") || (W18 == "") || (W19 == "") || (W20 == "") || (W21 == "") || (W22 == ""))
                                {
                                    MessageBox.Show("Thông tin của sản phẩm: " + G_CODE + "có thông tin trống, sẽ ko thêm PO này !");
                                }
                                else
                                {

                                    ProductBLL pro = new ProductBLL();
                                    string checkUSEYN = pro.checkM100UseYN(G_CODE);
                                    if (pro.checkFCSTExist(CUST_CD, G_CODE, YEAR, WEEKNO) != -1)
                                    {
                                        dataGridView1.Rows[row.Index].Cells[29].Value = "NG - Đã tồn tại FCST";
                                        dataGridView1.Rows[row.Index].Cells[29].Style.BackColor = Color.Red;
                                    }
                                    else if (checkUSEYN == "N")
                                    {
                                        dataGridView1.Rows[row.Index].Cells[29].Value = "NG - Ver này đã bị khoá, check lại ver final đi";
                                        dataGridView1.Rows[row.Index].Cells[29].Style.BackColor = Color.Red;
                                    }
                                    else
                                    {                                        
                                        dataGridView1.Rows[row.Index].Cells[29].Value = "OK";
                                        dataGridView1.Rows[row.Index].Cells[29].Style.BackColor = Color.LightGreen;
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

        }
        private void tạo1FCSTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //insertFCST(1);
        }

        private void checkFCSTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int total_flag = po_flag + invoice_flag + fcst_flag + khgh_flag;
            if (total_flag == 0)
            {

                //insertFCST(0);
            }
            else
            {
                MessageBox.Show("Check choác gì ! ?, import vào r mới check được !");
            }
        }

        public void xoaFCST()
        {
            if(fcst_flag==1)
            {
                if (textBox8.Text == "xoa")
                {

                    if (MessageBox.Show("Bạn thực sự muốn xóa FCST?", "Xóa FCST ?", MessageBoxButtons.OKCancel) == DialogResult.OK)
                    {
                        try
                        {
                            ProductBLL pro = new ProductBLL();
                            DataTable dt = new DataTable();

                            var selectedRows = dataGridView1.SelectedRows
                           .OfType<DataGridViewRow>()
                           .Where(row => !row.IsNewRow)
                           .ToArray();

                            progressBar1.Minimum = 0; //Đặt giá trị nhỏ nhất cho ProgressBar
                            progressBar1.Maximum = dataGridView1.SelectedRows.Count; //Đặt giá trị lớn nhất cho ProgressBar
                            int startprogress = 0;

                            foreach (var row in selectedRows)
                            {
                                string fcst_id = row.Cells["FCST_ID"].Value.ToString(); //cell 0
                                dt = pro.DeleteFCST(fcst_id);
                                pro.writeHistory("002", LoginID, "FCST TABLE", "XOA", "Xoa FCST", "" + fcst_id);
                                startprogress = startprogress + 1;
                                label4.Text = "Progress: " + startprogress + "/" + dataGridView1.SelectedRows.Count;
                                progressBar1.Value = startprogress;
                            }
                            progressBar1.Value = 0;
                            dataGridView1.ClearSelection();
                            MessageBox.Show("Xóa FCST thành công !");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.ToString());
                        }


                    }
                }
                else
                {
                    MessageBox.Show("Lêu Lêu, còn lâu mới xóa được nhé");
                }

            }
            else
            {
                MessageBox.Show("Không phải bảng FCST nên không xóa được!");
            }
            
        }
        private void xóaFCSTĐãChọnToolStripMenuItem_Click(object sender, EventArgs e)
        {
            xoaFCST();
        }

        private void tạo1KHGHMớiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int total_flag = po_flag + invoice_flag + fcst_flag + khgh_flag;
            if (total_flag == 0)
            {
                //insertKHGH(0);
            }
            else
            {
                MessageBox.Show("Check choác gì, import file vào mới check được !");

            }
        }
        public void xoaplan()
        {
            if(khgh_flag == 1)
            {
                if (textBox8.Text == "xoa")
                {
                    if (MessageBox.Show("Bạn thực sự muốn xóa PLAN?", "Xóa PLAN ?", MessageBoxButtons.OKCancel) == DialogResult.OK)
                    {
                        try
                        {
                            ProductBLL pro = new ProductBLL();
                            DataTable dt = new DataTable();

                            var selectedRows = dataGridView1.SelectedRows
                           .OfType<DataGridViewRow>()
                           .Where(row => !row.IsNewRow)
                           .ToArray();

                            progressBar1.Minimum = 0; //Đặt giá trị nhỏ nhất cho ProgressBar
                            progressBar1.Maximum = dataGridView1.SelectedRows.Count; //Đặt giá trị lớn nhất cho ProgressBar
                            int startprogress = 0;

                            foreach (var row in selectedRows)
                            {
                                string plan_id = row.Cells["PLAN_ID"].Value.ToString(); // cell 0
                                dt = pro.DeletePlan(plan_id);
                                pro.writeHistory("002", LoginID, "PLAN TABLE", "XOA", "Xoa PLAN", "" + plan_id);

                                startprogress = startprogress + 1;
                                label4.Text = "Progress: " + startprogress + "/" + dataGridView1.SelectedRows.Count;
                                progressBar1.Value = startprogress;
                            }
                            progressBar1.Value = 0;
                            dataGridView1.ClearSelection();
                            MessageBox.Show("Xóa PLAN thành công !");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.ToString());
                        }

                    }
                }
                else
                {
                    MessageBox.Show("Lêu lêu, còn lâu mới xóa được nhé !");
                }

            }
            else
            {
                MessageBox.Show("Không phải bảng plan nên  không xóa được");
            }
            
        }

        private void xóaPLANĐãChọnToolStripMenuItem_Click(object sender, EventArgs e)
        {
            xoaplan();              
        }

        private void thànhTíchGiaoHàngReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            reportForm RPF = new reportForm();
            RPF.Show();
        }

        private void pOBalanceReportToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void generalReport2ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            reportForm2 rpf2 = new reportForm2();
            rpf2.Show();
        }

        private void weekMonthReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            WeekMonthReportForm wmrp = new WeekMonthReportForm();
            wmrp.Show();
        }

        private void overdueReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form4 overdueform = new Form4();
            overdueform.Show();
        }

        private void wDeliveryPlanReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SOPForm sopf = new SOPForm();
            sopf.Show();
        }

       
        private void copyTableWithHeaderToolStripMenuItem_Click_1(object sender, EventArgs e)
        {            
            CopyToClipboardWithHeaders(dataGridView1);
        }

        private void checkCodeThiếuDataToolStripMenuItem_Click(object sender, EventArgs e)
        {
            po_flag = 0;
            invoice_flag = 0;
            khgh_flag = 0;
            fcst_flag = 0;
            ycsx_flag = 0;

            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dt = pro.report_NoData();
            this.dataGridView1.DataSource = null;
            this.dataGridView1.Rows.Clear();
            this.dataGridView1.DataSource = dt;
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
                            CUST_CD = Convert.ToString(row.Cells["CUST_CD"].Value);
                            EMPL_NO = Convert.ToString(row.Cells["EMPL_NO"].Value);
                            G_CODE = Convert.ToString(row.Cells["G_CODE"].Value);
                            PO_NO = Convert.ToString(row.Cells["PO_NO"].Value);
                            DELIVERY_QTY = Convert.ToString(row.Cells["DELIVERY_QTY"].Value);
                            DELIVERY_DATE = Convert.ToString(row.Cells["DELIVERY_DATE"].Value);
                            NOCANCEL = Convert.ToString(row.Cells["NOCANCEL"].Value);
                            INVOICE_NO = Convert.ToString(row.Cells["INVOICE_NO"].Value);


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
        private void upINVOICENOToolStripMenuItem_Click(object sender, EventArgs e)
        {
            updateINVOICENO();
        }

        private void lấyListCodeCMSToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.Show();
            ProductBLL pro = new ProductBLL();
            DataTable dt = pro.report_getCodeList();
            this.dataGridView1.DataSource = null;
            this.dataGridView1.Rows.Clear();
            this.dataGridView1.DataSource = dt;
        }

        private void lấyListKháchHàngToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.Show();
            ProductBLL pro = new ProductBLL();
            DataTable dt = pro.report_getCustomerList();
            this.dataGridView1.DataSource = null;
            this.dataGridView1.Rows.Clear();
            this.dataGridView1.DataSource = dt;
        }

        private void lấyListNhânViênToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.Show();
            dataGridView1.Show();
            ProductBLL pro = new ProductBLL();
            DataTable dt = pro.report_getEmployeeList();
            this.dataGridView1.DataSource = null;
            this.dataGridView1.Rows.Clear();
            this.dataGridView1.DataSource = dt;
        }

        private void newPOToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //dataGridView1.ReadOnly = false;
            //dataGridView1.DataSource = null;
            //dataGridView1.Rows.Clear();

            
        }

        private void Form1_Click(object sender, EventArgs e)
        {

        }

        public static DateTime FirstDayOfWeek(DateTime date)
        {
            DayOfWeek fdow = CultureInfo.CurrentCulture.DateTimeFormat.FirstDayOfWeek;
            int offset = fdow - date.DayOfWeek;
            DateTime fdowDate = date.AddDays(offset);
            return fdowDate;
        }

        public static DateTime LastDayOfWeek(DateTime date)
        {
            DateTime ldowDate = FirstDayOfWeek(date).AddDays(6);
            return ldowDate;
        }

        private void báoGiáToolStripMenuItem_Click(object sender, EventArgs e)
        {
            /*
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            String ngaygiohethong = pro.getsystemDateTime();
            String ngayhethong = ngaygiohethong.Substring(0,4) + ngaygiohethong.Substring(5,2) + ngaygiohethong.Substring(8,2);
            MessageBox.Show(ngayhethong);
            */         
            
        }

        public void suainvoice()
        {
            if (invoice_flag == 1)
            {
                InvoiceForm invoiceform = new InvoiceForm();
                invoiceform.loginIDInvoiceForm = LoginID;

                var selectedRows = dataGridView1.SelectedRows
                      .OfType<DataGridViewRow>()
                      .Where(row => !row.IsNewRow)
                      .ToArray();
                foreach (var row in selectedRows)
                {
                    try
                    {
                        string CUST_CD, EMPL_NO, G_CODE, PO_NO, DELIVERY_QTY, DELIVERY_DATE, NOCANCEL, REMARK, G_NAME, CUST_NAME, DELIVERY_ID;
                        string CTR_CD = "002";
                        G_CODE = row.Cells[0].Value.ToString();
                        DELIVERY_QTY = row.Cells[6].Value.ToString(); ;
                        CUST_CD = row.Cells[17].Value.ToString();
                        EMPL_NO = row.Cells[18].Value.ToString();
                        //DELIVERY_DATE = STYMD(DateTime.Today.Year, DateTime.Today.Month, DateTime.Today.Day); 
                        DELIVERY_DATE = row.Cells[3].Value.ToString().Substring(0, 10);
                        PO_NO = row.Cells[8].Value.ToString();
                        REMARK = row.Cells[14].Value.ToString();
                        G_NAME = row.Cells[4].Value.ToString();
                        CUST_NAME = row.Cells[2].Value.ToString();
                        DELIVERY_ID = row.Cells[15].Value.ToString();

                        NOCANCEL = "1";

                        invoiceform.CUST_CD = CUST_CD;
                        invoiceform.EMPL_NO = EMPL_NO;
                        invoiceform.G_CODE = G_CODE;
                        invoiceform.PO_NO = PO_NO;
                        invoiceform.DELIVERY_QTY = DELIVERY_QTY;
                        invoiceform.DELIVERY_DATE = DELIVERY_DATE;
                        invoiceform.NOCANCEL = NOCANCEL;
                        invoiceform.G_NAME = G_NAME;
                        invoiceform.CUST_NAME = CUST_NAME;
                        invoiceform.DELIVERY_ID = DELIVERY_ID;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }

                }
                invoiceform.updateform();
                invoiceform.Show();
            }
            else
            {
                MessageBox.Show("Dữ liệu hiện tại k fai bảng INVOICE, không sửa được !");
            }
        }
        private void sửaInvoiceToolStripMenuItem_Click(object sender, EventArgs e)
        {           
                suainvoice();           
        }

        private void xóaInvoiceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            xoaInvoice();
        }


        public void suaPO()
        {
            if (po_flag == 1)
            {
                POForm poForm = new POForm();
                poForm.loginIDpoForm = LoginID;

                var selectedRows = dataGridView1.SelectedRows
                      .OfType<DataGridViewRow>()
                      .Where(row => !row.IsNewRow)
                      .ToArray();
                foreach (var row in selectedRows)
                {
                    try
                    {
                        string CUST_CD, EMPL_NO, G_CODE, PO_NO, PO_DATE, RD_DATE, PROD_PRICE, DELIVERY_QTY, DELIVERY_DATE, REMARK, G_NAME, CUST_NAME, PO_QTY, PO_ID;
                        string CTR_CD = "002";
                        G_CODE = row.Cells["G_CODE"].Value.ToString();
                        CUST_CD = row.Cells["CUST_CD"].Value.ToString();
                        EMPL_NO = row.Cells["EMPL_NO"].Value.ToString();
                        //DELIVERY_DATE = STYMD(DateTime.Today.Year, DateTime.Today.Month, DateTime.Today.Day); 
                        //DELIVERY_DATE = row.Cells[3].Value.ToString().Substring(0, 10);
                        PO_NO = row.Cells["PO_NO"].Value.ToString();  //ok
                        REMARK = row.Cells["REMARK"].Value.ToString(); //ok
                        G_NAME = row.Cells["G_NAME"].Value.ToString(); //ok
                        CUST_NAME = row.Cells["CUST_NAME_KD"].Value.ToString();//ok
                        PO_QTY = row.Cells["PO_QTY"].Value.ToString();
                        PO_DATE = row.Cells["PO_DATE"].Value.ToString().ToString().Substring(0, 10);
                        RD_DATE = row.Cells["RD_DATE"].Value.ToString().ToString().Substring(0, 10);
                        PROD_PRICE = row.Cells["PROD_PRICE"].Value.ToString();
                        PO_ID = row.Cells["PO_ID"].Value.ToString();  //ok


                        poForm.CUST_CD = CUST_CD;
                        poForm.EMPL_NO = EMPL_NO;
                        poForm.G_CODE = G_CODE;
                        poForm.PO_NO = PO_NO;
                        poForm.PO_QTY = PO_QTY;
                        poForm.PO_DATE = PO_DATE;
                        poForm.RD_DATE = RD_DATE;
                        poForm.PROD_PRICE = PROD_PRICE;
                        poForm.PO_ID = PO_ID;

                        poForm.G_NAME = G_NAME;
                        poForm.CUST_NAME = CUST_NAME;


                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }

                }
                poForm.updateform();
                poForm.Show();
            }
            else
            {
                MessageBox.Show("Dữ liệu hiện tại k fai bảng PO, không sửa được !");
            }
        }
        private void sửaPOToolStripMenuItem_Click(object sender, EventArgs e)
        {            
                suaPO();
        }

        private void xóaPOToolStripMenuItem_Click(object sender, EventArgs e)
        {
                xoaPO();                      
        }

        public void traBOM()
        {

            bom_flag = 1;
            fcst_flag = 0;
            invoice_flag = 0;
            khgh_flag = 0;
            po_flag = 0;
            ycsx_flag = 0;

            dataGridView1.Show();
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dt = pro.info_getCODEInfo(textBox1.Text);
            this.dataGridView1.DataSource = null;
            this.dataGridView1.Columns.Clear();
            dataGridView1.DataSource = dt;
            if (dt.Rows.Count > 0)
            {
                MessageBox.Show("Đã load " + dt.Rows.Count + " dòng, double click vào code để xem BOM");

            }
            else
            {
                MessageBox.Show("Không có code này !");
            }
            po_flag = 0;
            invoice_flag = 0;
            khgh_flag = 0;
            fcst_flag = 0;
            ycsx_flag = 0;
            bom_flag = 1;

        }
        public void traBOM_Async(DoWorkEventArgs e)
        {
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dtgv1_data = pro.info_getCODEInfo(textBox1.Text);           

        }

        private void button27_Click(object sender, EventArgs e)
        {
            traBOM();
        }

        private void xóaLanToolStripMenuItem_Click(object sender, EventArgs e)
        {
            xoaplan();
        }

        private void xóaFCSTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            xoaFCST();
        }

        private void checkUploadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ProductBLL pro = new ProductBLL();
            DataTable dt = pro.readHistory();
            dataGridView1.DataSource = dt;            
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {            
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop, false);

            foreach(string file in files)
            { 
                po_flag = 0;
                invoice_flag = 0;
                khgh_flag = 0;
                fcst_flag = 0;
                ycsx_flag = 0;
                try
                {
                    dataGridView1.Columns.Clear();
                    if (file != "")
                    {
                        dsNV = ExcelFactory.readFromExcelFile(file);
                        this.dataGridView1.DataSource = null;
                        this.dataGridView1.Columns.Clear();
                        dataGridView1.DataSource = dsNV;
                        button4.Enabled = true;
                        button3.Enabled = true;
                        button16.Enabled = true;
                        import_excel_flag = 1;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
                MessageBox.Show("Import file hoàn thành!");
            }
        }

        private void nhậpPOToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
        }

        private void nhậpTrựcTiếpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            InputForm inputForm = new InputForm();
            inputForm.LoginID = LoginID;
            inputForm.Show();
        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.S)
            {
                if (dataGridView1 == null || dataGridView1.Rows.Count == 0)
                {
                    MessageBox.Show("Không có dữ liệu!");
                }
                else
                {
                    ExcelFactory.writeToExcelFile(dataGridView1);
                }
            }

            else if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.O)
            {
                po_flag = 0;
                invoice_flag = 0;
                khgh_flag = 0;
                fcst_flag = 0;
                ycsx_flag = 0;
                try
                {
                    OpenFileDialog ofd = new OpenFileDialog();
                    ofd.ShowDialog();
                    string file = ofd.FileName;

                    if (file != "")
                    {
                        dsNV = ExcelFactory.readFromExcelFile(file);
                        this.dataGridView1.DataSource = null;
                        this.dataGridView1.Columns.Clear();
                        dataGridView1.DataSource = dsNV;
                        button4.Enabled = true;
                        button3.Enabled = true;
                        button16.Enabled = true;
                        /*
                        dataGridView1.Columns[0].HeaderText = "Mã Nhân Viên";
                        dataGridView1.Columns[1].HeaderText = "Tên Nhân Viên";
                        dataGridView1.Columns[2].HeaderText = "Tuổi Nhân Viên";
                        */

                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
            else if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.P)
            {
                traPO();
            }
            else if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.D)
            {
                traInvoice();
            }
            else if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.F)
            {
                traFCST();
            }
            else if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.L)
            {
                traPLAN();
            }
            else if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.N)
            {
                InputForm inputForm = new InputForm();
                inputForm.LoginID = LoginID;
                inputForm.Show();
            }
            else if (Control.ModifierKeys == Keys.Shift && e.KeyCode == Keys.P)
            {
                POForm poform = new POForm();
                poform.loginIDpoForm = LoginID;
                poform.Show();
            }
            else if (Control.ModifierKeys == Keys.Shift && e.KeyCode == Keys.D)
            {
                InvoiceForm invoiceform = new InvoiceForm();
                invoiceform.loginIDInvoiceForm = LoginID;
                invoiceform.Show();
            }
            else if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.E)
            {

                if (dataGridView1.SelectedCells.Count > 0)
                {
                    if(dataGridView1.ReadOnly == true)
                    {
                        dataGridView1.ReadOnly = false;
                        this.dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
                        dataGridView1.SelectionMode = DataGridViewSelectionMode.CellSelect;
                        MessageBox.Show("Bật sửa");
                    }
                    else if(dataGridView1.ReadOnly == false)
                    {
                        dataGridView1.ReadOnly = true;
                        this.dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
                        dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                        MessageBox.Show("Tắt sửa");
                    }

                }

                
            }
        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.S)
            {
                if (dataGridView1 == null || dataGridView1.Rows.Count == 0)
                {
                    MessageBox.Show("Không có dữ liệu!");
                }
                else
                {
                    ExcelFactory.writeToExcelFile(dataGridView1);
                }
            }

            else if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.O)
            {
                po_flag = 0;
                invoice_flag = 0;
                khgh_flag = 0;
                fcst_flag = 0;
                ycsx_flag = 0;
                try
                {
                    OpenFileDialog ofd = new OpenFileDialog();
                    ofd.ShowDialog();
                    string file = ofd.FileName;

                    if (file != "")
                    {
                        dsNV = ExcelFactory.readFromExcelFile(file);
                        this.dataGridView1.DataSource = null;
                        this.dataGridView1.Columns.Clear();
                        dataGridView1.DataSource = dsNV;
                        button4.Enabled = true;
                        button3.Enabled = true;
                        button16.Enabled = true;
                        /*
                        dataGridView1.Columns[0].HeaderText = "Mã Nhân Viên";
                        dataGridView1.Columns[1].HeaderText = "Tên Nhân Viên";
                        dataGridView1.Columns[2].HeaderText = "Tuổi Nhân Viên";
                        */

                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
            else if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.P)
            {
                traPO();
            }
            else if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.D)
            {
                traInvoice();
            }
            else if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.F)
            {
                traFCST();
            }
            else if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.L)
            {
                traPLAN();
            }
            else if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.N)
            {
                InputForm inputForm = new InputForm();
                inputForm.LoginID = LoginID;
                inputForm.Show();
            }
            else if (Control.ModifierKeys == Keys.Control && Control.ModifierKeys == Keys.Shift &&  e.KeyCode == Keys.P)
            {
                POForm poform = new POForm();
                poform.loginIDpoForm = LoginID;
                poform.Show();
            }
            else if (Control.ModifierKeys == Keys.Control && Control.ModifierKeys == Keys.Shift && e.KeyCode == Keys.D)
            {
                InvoiceForm invoiceform = new InvoiceForm();
                invoiceform.loginIDInvoiceForm = LoginID;
                invoiceform.Show();
            }

        }
        public int GetWeekNumber(string dt)
        {
            DateTime dd = DateTime.Parse(dt);
            dd = dd.AddDays(1);
            CultureInfo ciCurr = CultureInfo.CurrentCulture;
            int weekNum = ciCurr.Calendar.GetWeekOfYear(dd, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
            return weekNum;
        }

        private void uploadFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string year = DateTime.Now.Year.ToString();
            MessageBox.Show(""+GetWeekNumber(DateTime.Now.Year.ToString() + "-12-31"));

        }

        private void traNhậpXuấtKiểmToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void kiểmTraTiếnĐộSXToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var selectedRows = dataGridView1.SelectedRows
                         .OfType<DataGridViewRow>()
                         .Where(row => !row.IsNewRow)
                         .ToArray();
            foreach (var row in selectedRows)
            {
                string codecms = row.Cells["G_CODE"].Value.ToString();
                //MessageBox.Show(ycsxno);
                YCSX_Manager ycsxmanager = new YCSX_Manager();
                ycsxmanager.Show();
                ycsxmanager.tratinhhinh(codecms);
                MessageBox.Show("Sẽ show các ycsx còn pending");
               
            }



        }

        private void quảnLýYCSXToolStripMenuItem_Click(object sender, EventArgs e)
        {
            YCSX_Manager ycsxmanager = new YCSX_Manager();
            ycsxmanager.Login_ID = LoginID;
            ycsxmanager.Show();
        }

        private void cậpNhậtPhầnMềmToolStripMenuItem_Click(object sender, EventArgs e)
        {
           /*System.Diagnostics.Process.Start("http://14.160.33.94/update/ERP2/lastest.zip"); */
        }

        private void traNhậpXuấtKiểmToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Form inspectionform = new INSPECTION();
            inspectionform.Show();
        }

        private void thêmGiaoHàngChoPOToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if(po_flag == 1)
            {
                InvoiceForm invoiceform = new InvoiceForm();
                invoiceform.loginIDInvoiceForm = LoginID;

                var selectedRows = dataGridView1.SelectedRows
                      .OfType<DataGridViewRow>()
                      .Where(row => !row.IsNewRow)
                      .ToArray();
                foreach (var row in selectedRows)
                {
                    try
                    {
                        string CUST_CD, EMPL_NO, G_CODE, PO_NO, DELIVERY_QTY, DELIVERY_DATE, NOCANCEL, REMARK, G_NAME, CUST_NAME;
                        string CTR_CD = "002";
                        G_CODE = row.Cells["G_CODE"].Value.ToString();
                        DELIVERY_QTY = "";
                        CUST_CD = row.Cells["CUST_CD"].Value.ToString();
                        EMPL_NO = row.Cells["EMPL_NO"].Value.ToString();
                        //DELIVERY_DATE = STYMD(DateTime.Today.Year, DateTime.Today.Month, DateTime.Today.Day); 
                        DateTime today = DateTime.Today;
                        today = today.AddDays(-1);
                        DELIVERY_DATE = today.ToString("yyyy-MM-dd");
                        PO_NO = row.Cells["PO_NO"].Value.ToString();
                        REMARK = "";
                        G_NAME = row.Cells["G_NAME"].Value.ToString();
                        CUST_NAME = row.Cells["CUST_NAME_KD"].Value.ToString();
                        NOCANCEL = "1";

                        invoiceform.CUST_CD = CUST_CD;
                        invoiceform.EMPL_NO = EMPL_NO;
                        invoiceform.G_CODE = G_CODE;
                        invoiceform.PO_NO = PO_NO;
                        invoiceform.DELIVERY_QTY = DELIVERY_QTY;
                        invoiceform.DELIVERY_DATE = DELIVERY_DATE;
                        invoiceform.NOCANCEL = NOCANCEL;
                        invoiceform.G_NAME = G_NAME;
                        invoiceform.CUST_NAME = CUST_NAME;


                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }


                }
                invoiceform.updateform();
                invoiceform.Show();
            }
            else
            {
                MessageBox.Show("Không phải bảng PO, ko thêm được giao hàng");
            }
            
        }

        private void lấyBOMToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void taInvoiceToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void traFCSTToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void traKHGHToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void upThôngTin1CodeToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void updateSốHóaĐơnToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Đang phát triển tính năng");            

        }

        private void uPThôngTinCodeQLSXToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
        }

        private void checkCodeThiếuDataQLSXToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
        }

        private void checkTínhĐúngĐắnCủaDataCapaToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
        }

        private void checkCodeThiếuDataQLSXToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            po_flag = 0;
            invoice_flag = 0;
            khgh_flag = 0;
            fcst_flag = 0;
            ycsx_flag = 0;
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dt = pro.report_QLSX();
            this.dataGridView1.DataSource = null;
            this.dataGridView1.Rows.Clear();
            this.dataGridView1.DataSource = dt;
            MessageBox.Show("Đã load : " + dt.Rows.Count + " dòng");
            MessageBox.Show("Đã load danh sách các code còn tồn PO, có xuất hiện trong kế hoạch giao hàng SOP 14 ngày trở lại đây, có xuất hiện trong phòng kiểm tra 14 ngày trở lại đây. Mà vẫn chưa được update thông tin");
        }

        private void checkQuyTắcCapaQLSXToolStripMenuItem_Click(object sender, EventArgs e)
        {
            po_flag = 0;
            invoice_flag = 0;
            khgh_flag = 0;
            fcst_flag = 0;
            ycsx_flag = 0;
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dt = pro.report_QLSX_validating();
            this.dataGridView1.DataSource = null;
            this.dataGridView1.Rows.Clear();
            this.dataGridView1.DataSource = dt;
            MessageBox.Show("Đã load : " + dt.Rows.Count + " dòng");
            MessageBox.Show("Đã load danh sách các code điền ko đúng quy tắc");
        }

        private void updateDatadòngĐcChọnToolStripMenuItem_Click(object sender, EventArgs e)
        {
            upthongtincodeQLSXhangloat();
        }

        public void checkBTP()
        {
            po_flag = 0;
            invoice_flag = 0;
            khgh_flag = 0;
            fcst_flag = 0;
            ycsx_flag = 0;
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            string fromdate, todate;
            fromdate = STYMD2(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day);
            
            dt = pro.report_checkBTP(generate_condition_BTP());
            this.dataGridView1.DataSource = null;
            this.dataGridView1.Rows.Clear();
            this.dataGridView1.DataSource = dt;
            formatDataGridViewBTP(dataGridView1);
            MessageBox.Show("Đã load : " + dt.Rows.Count + " dòng, BTP ngày " + fromdate);            

        }

        public void checkBTP2()
        {
            po_flag = 0;
            invoice_flag = 0;
            khgh_flag = 0;
            fcst_flag = 0;
            ycsx_flag = 0;
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            string fromdate, todate;
            fromdate = STYMD2(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day);

            dt = pro.report_checkBTP2(generate_condition_BTP());
            this.dataGridView1.DataSource = null;
            this.dataGridView1.Rows.Clear();
            this.dataGridView1.DataSource = dt;
            formatDataGridViewBTP(dataGridView1);
            MessageBox.Show("Đã load : " + dt.Rows.Count + " dòng, BTP ngày " + fromdate);

        }


        public void checkCK()
        {
            po_flag = 0;
            invoice_flag = 0;
            khgh_flag = 0;
            fcst_flag = 0;
            ycsx_flag = 0;
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            string fromdate, todate;
            fromdate = STYMD2(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day);
            
            dt = pro.report_checkCK(generate_condition_chokiem());
            this.dataGridView1.DataSource = null;
            this.dataGridView1.Rows.Clear();
            this.dataGridView1.DataSource = dt;
            formatDataGridViewChoKiem(dataGridView1);
            MessageBox.Show("Đã load : " + dt.Rows.Count + " dòng");

        }

        public void checkCK2()
        {
            po_flag = 0;
            invoice_flag = 0;
            khgh_flag = 0;
            fcst_flag = 0;
            ycsx_flag = 0;
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            string fromdate, todate;
            fromdate = STYMD2(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day);

            dt = pro.report_checkCK2(generate_condition_chokiem());
            this.dataGridView1.DataSource = null;
            this.dataGridView1.Rows.Clear();
            this.dataGridView1.DataSource = dt;
            formatDataGridViewChoKiem(dataGridView1);
            MessageBox.Show("Đã load : " + dt.Rows.Count + " dòng");

        }


        private void button29_Click(object sender, EventArgs e)
        {
            dataGridView1.Columns.Clear();
            checkBTP();
        }

        private void button30_Click(object sender, EventArgs e)
        {
            dataGridView1.Columns.Clear();
            checkCK();
        }
        public int hideshow = 0;
        private void button31_Click(object sender, EventArgs e)
        {
         
            
            
        }

        private void button28_Click(object sender, EventArgs e)
        {
            KHOTHANHPHAM ktp = new KHOTHANHPHAM();
            ktp.Show();
        }

        private void button31_Click_1(object sender, EventArgs e)
        {
            MessageBox.Show("Đang phát triển");
        }

        private void button33_Click(object sender, EventArgs e)
        {
            dataGridView1.Columns.Clear();
            checkCK2();
        }

        private void button32_Click(object sender, EventArgs e)
        {
            dataGridView1.Columns.Clear();
            checkBTP2();
        }

        private void newCodeBOMToolStripMenuItem_Click(object sender, EventArgs e)
        {
            NewCodeBom newCodeBom = new NewCodeBom();
            newCodeBom.EMPL_NO = LoginID;
            newCodeBom.Show();
        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void thôngTinVậtLiệuToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MaterialInfo mt = new MaterialInfo();
            mt.EMPL_NO = LoginID;
            mt.Show();
        }

        private void kiểmTraDataBOMGiáToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CheckBOMGia ckbom = new CheckBOMGia();
            ckbom.Show();
        }

        private void tínhBáoGiáToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TinhBaoGia tbg = new TinhBaoGia();
            tbg.Show();
        }

        private void tìnhHìnhSXTheoĐầuMáyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ProductionPlan prdtpl = new ProductionPlan();
            prdtpl.Show();
        }

        private void stockLiệuToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        public string process_lot_no_generate(string machine_name)
        {

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
            string LOT_HEADER = machine_name + new Form1().CreateHeader2();
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

        private void button34_Click(object sender, EventArgs e)
        {
            //MessageBox.Show(process_lot_no_generate("SR"));
            /*if(pictureBox1.Visible == true)
            {
                pictureBox1.Hide();
            }
            else
            {
                pictureBox1.Show();
            }
            */
            if(!backgroundWorker1.IsBusy)
            {
                
                backgroundWorker1.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Đang chạy tiến trình khác, đợi tiến trình đó chạy xong đã rồi thử lại");
            }           

        }

        private void sXToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void thêmKháchMớiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            QuanLyKhachHang ql = new QuanLyKhachHang();
            ql.Show();
        }

        private void yêuCầuSảnXuấtToolStripMenuItem_Click(object sender, EventArgs e)
        {
            YCSX_Manager ycsxmanager = new YCSX_Manager();
            ycsxmanager.Login_ID = LoginID;
            ycsxmanager.Show();
        }

        private void tạo1POMơiToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            POForm poform = new POForm();
            poform.loginIDpoForm = LoginID;
            poform.Show();
        }

        private void checkPOToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            checkPO();
        }

        private void uploadHàngLoạtToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            upPOhangloat();
        }

        private void newPOToolStripMenuItem_Click_1(object sender, EventArgs e)
        {

        }

        private void tạo1InvoiceToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            InvoiceForm invoiceform = new InvoiceForm();
            invoiceform.loginIDInvoiceForm = LoginID;
            invoiceform.Show();
        }

        private void checkInvoiceToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            checkInvoice();
        }

        private void uploadInvoiceHàngLoạtToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            upInvoicehangloat();
        }

        private void upINVOICENOToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            updateINVOICENO();
        }

        private void checkFCSTToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            int total_flag = po_flag + invoice_flag + fcst_flag + khgh_flag;
            if (total_flag == 0)
            {

                //insertFCST(0);
            }
            else
            {
                MessageBox.Show("Check choác gì ! ?, import vào r mới check được !");
            }
        }

        private void tạo1FCSTToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            //insertFCST(1);
        }

        private void uploadKHGHHàngLoạtToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            //insertKHGH(1);
        }

        private void tạo1KHGHMớiToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            int total_flag = po_flag + invoice_flag + fcst_flag + khgh_flag;
            if (total_flag == 0)
            {
                //insertKHGH(0);
            }
            else
            {
                MessageBox.Show("Check choác gì, import file vào mới check được !");

            }
        }

        private void thànhTíchGiaoHàngReportToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            reportForm RPF = new reportForm();
            RPF.Show();
        }

        private void generalReport2ToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            reportForm2 rpf2 = new reportForm2();
            rpf2.Show();
        }

        private void overdueReportToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            Form4 overdueform = new Form4();
            overdueform.Show();
        }

        private void wDeliveryPlanReportToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            SOPForm sopf = new SOPForm();
            sopf.Show();
        }

        private void weekMonthReportToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            WeekMonthReportForm wmrp = new WeekMonthReportForm();
            wmrp.Show();
        }

        private void checkUploadToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            ProductBLL pro = new ProductBLL();
            DataTable dt = pro.readHistory();
            dataGridView1.DataSource = dt;
        }

        private void lấyListCodeCMSToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            dataGridView1.Show();
            ProductBLL pro = new ProductBLL();
            DataTable dt = pro.report_getCodeList();
            this.dataGridView1.DataSource = null;
            this.dataGridView1.Rows.Clear();
            this.dataGridView1.DataSource = dt;
        }

        private void lấyListKháchHàngToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            dataGridView1.Show();
            ProductBLL pro = new ProductBLL();
            DataTable dt = pro.report_getCustomerList();
            this.dataGridView1.DataSource = null;
            this.dataGridView1.Rows.Clear();
            this.dataGridView1.DataSource = dt;
        }

        private void lấyListNhânViênToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            dataGridView1.Show();
            dataGridView1.Show();
            ProductBLL pro = new ProductBLL();
            DataTable dt = pro.report_getEmployeeList();
            this.dataGridView1.DataSource = null;
            this.dataGridView1.Rows.Clear();
            this.dataGridView1.DataSource = dt;
        }

        private void thêmKháchMớiToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            QuanLyKhachHang ql = new QuanLyKhachHang();
            ql.Show();
        }

        private void upThôngTinHàngLoạtToolStripMenuItem1_Click_1(object sender, EventArgs e)
        {
            uploadcustomerinfor();
        }

        private void upThôngTinHàngLoạtToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            updatecodeinfo();
        }

        private void checkCodeThiếuDataToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            po_flag = 0;
            invoice_flag = 0;
            khgh_flag = 0;
            fcst_flag = 0;
            ycsx_flag = 0;

            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dt = pro.report_NoData();
            this.dataGridView1.DataSource = null;
            this.dataGridView1.Rows.Clear();
            this.dataGridView1.DataSource = dt;
        }

        private void upThôngTinCODEToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void checkCodeThiếuDataQLSXToolStripMenuItem1_Click_1(object sender, EventArgs e)
        {
            po_flag = 0;
            invoice_flag = 0;
            khgh_flag = 0;
            fcst_flag = 0;
            ycsx_flag = 0;
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dt = pro.report_QLSX();
            dataGridView1.Columns.Clear();
            this.dataGridView1.DataSource = null;
            this.dataGridView1.Rows.Clear();
            this.dataGridView1.DataSource = dt;
            MessageBox.Show("Đã load : " + dt.Rows.Count + " dòng");
            MessageBox.Show("Đã load danh sách các code còn tồn PO, có xuất hiện trong kế hoạch giao hàng SOP 14 ngày trở lại đây, có xuất hiện trong phòng kiểm tra 14 ngày trở lại đây. Mà vẫn chưa được update thông tin");
        }

        private void checkQuyTắcCapaQLSXToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            po_flag = 0;
            invoice_flag = 0;
            khgh_flag = 0;
            fcst_flag = 0;
            ycsx_flag = 0;
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dt = pro.report_QLSX_validating();
            dataGridView1.Columns.Clear();
            this.dataGridView1.DataSource = null;
            this.dataGridView1.Rows.Clear();
            this.dataGridView1.DataSource = dt;
            MessageBox.Show("Đã load : " + dt.Rows.Count + " dòng");
            MessageBox.Show("Đã load danh sách các code điền ko đúng quy tắc");
        }

        private void updateDatadòngĐcChọnToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            upthongtincodeQLSXhangloat();
        }

        private void checkPOToolStripMenuItem1_Click(object sender, EventArgs e)
        {            
            po_flag = 0;
            bom_flag = 0;
            fcst_flag = 0;
            invoice_flag = 0;
            ycsx_flag = 0;
            khgh_flag = 0;

            
            check_po_flag = 1;
            up_po_flag = 0;
            check_invoice_flag = 0;
            up_invoice_flag = 0;
            check_fcst_flag = 0;
            up_fcst_flag = 0;
            check_plan_flag = 0;
            up_plan_flag = 0;
            if (!backgroundWorker1.IsBusy)
            {
                backgroundWorker1.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Tiến trình khác đang chạy, thử lại sau");
            }

        }

        private void pOToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            POForm poform = new POForm();
            poform.loginIDpoForm = LoginID;
            poform.Show();
        }

        private void nhiềuPOToolStripMenuItem_Click(object sender, EventArgs e)
        {            
            po_flag = 0;
            bom_flag = 0;
            fcst_flag = 0;
            invoice_flag = 0;
            ycsx_flag = 0;
            khgh_flag = 0;

            import_excel_flag = 0;
            check_po_flag = 0;
            up_po_flag = 1;
            check_invoice_flag = 0;
            up_invoice_flag = 0;
            check_fcst_flag = 0;
            up_fcst_flag = 0;
            check_plan_flag = 0;
            up_plan_flag = 0;
            if (!backgroundWorker1.IsBusy)
            {
                backgroundWorker1.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Tiến trình khác đang chạy, thử lại sau");
            }
        }

        private void invoiceToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            InvoiceForm invoiceform = new InvoiceForm();
            invoiceform.loginIDInvoiceForm = LoginID;
            invoiceform.Show();
        }

        private void checkInvoiceToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            //checkInvoice();
            po_flag = 0;
            bom_flag = 0;
            fcst_flag = 0;
            invoice_flag = 0;
            ycsx_flag = 0;
            khgh_flag = 0;

            
            check_po_flag = 0;
            up_po_flag = 0;
            check_invoice_flag = 1;
            up_invoice_flag = 0;
            check_fcst_flag = 0;
            up_fcst_flag = 0;
            check_plan_flag = 0;
            up_plan_flag = 0;
            if (!backgroundWorker1.IsBusy)
            {
                backgroundWorker1.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Tiến trình khác đang chạy, thử lại sau");
            }
        }

        private void upInvoiceHàngLoạtToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //upInvoicehangloat();
            po_flag = 0;
            bom_flag = 0;
            fcst_flag = 0;
            invoice_flag = 0;
            ycsx_flag = 0;
            khgh_flag = 0;

            import_excel_flag = 0;
            check_po_flag = 0;
            up_po_flag = 0;
            check_invoice_flag = 0;
            up_invoice_flag = 1;
            check_fcst_flag = 0;
            up_fcst_flag = 0;
            check_plan_flag = 0;
            up_plan_flag = 0;
            if (!backgroundWorker1.IsBusy)
            {
                backgroundWorker1.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Tiến trình khác đang chạy, thử lại sau");
            }
        }

        private void upInvoiceNoToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            updateINVOICENO();
        }

        public void checkFCST_Async(DoWorkEventArgs e)
        {
            
            if (import_excel_flag == 1)
            {

                insertFCST(0,e);
            }
            else
            {
                MessageBox.Show("Check choác gì ! ?, import vào r mới check được !");
            }
        }
        private void checkFCSTToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            po_flag = 0;
            bom_flag = 0;
            fcst_flag = 0;
            invoice_flag = 0;
            ycsx_flag = 0;
            khgh_flag = 0;

            
            check_po_flag = 0;
            up_po_flag = 0;
            check_invoice_flag = 0;
            up_invoice_flag = 0;
            check_fcst_flag = 1;
            up_fcst_flag = 0;
            check_plan_flag = 0;
            up_plan_flag = 0;
            if (!backgroundWorker1.IsBusy)
            {
                backgroundWorker1.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Tiến trình khác đang chạy, thử lại sau");
            }


        }

        public void uploadFCST_Async(DoWorkEventArgs e)
        {
            insertFCST(1,e);
        }
        private void upFCSTHàngLoạtToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //insertFCST(1);
            po_flag = 0;
            bom_flag = 0;
            fcst_flag = 0;
            invoice_flag = 0;
            ycsx_flag = 0;
            khgh_flag = 0;

            import_excel_flag = 0;
            check_po_flag = 0;
            up_po_flag = 0;
            check_invoice_flag = 0;
            up_invoice_flag = 0;
            check_fcst_flag = 0;
            up_fcst_flag = 1;
            check_plan_flag = 0;
            up_plan_flag = 0;
            if (!backgroundWorker1.IsBusy)
            {
                backgroundWorker1.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Tiến trình khác đang chạy, thử lại sau");
            }
        }

        public void checkPLAN_Async(DoWorkEventArgs e)
        {
           // MessageBox.Show("import_excel_flag=" + import_excel_flag);
            if (import_excel_flag == 1)
            {
                insertKHGH(0,e);
            }
            else
            {
                MessageBox.Show("Check choác gì, import file vào mới check được !");

            }
        }
        private void checkPLANToolStripMenuItem_Click(object sender, EventArgs e)
        {
            po_flag = 0;
            bom_flag = 0;
            fcst_flag = 0;
            invoice_flag = 0;
            ycsx_flag = 0;
            khgh_flag = 0;
            
            check_po_flag = 0;
            up_po_flag = 0;
            check_invoice_flag = 0;
            up_invoice_flag = 0;
            check_fcst_flag = 0;
            up_fcst_flag = 0;
            check_plan_flag = 1;
            up_plan_flag = 0;
            if (!backgroundWorker1.IsBusy)
            {
                backgroundWorker1.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Tiến trình khác đang chạy, thử lại sau");
            }
        }

        public void upPLAN_Async(DoWorkEventArgs e)
        {
            insertKHGH(1,e);
        }
        private void upPlanHàngLoạtToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //insertKHGH(1);
            po_flag = 0;
            bom_flag = 0;
            fcst_flag = 0;
            invoice_flag = 0;
            ycsx_flag = 0;
            khgh_flag = 0;

            import_excel_flag = 0;
            check_po_flag = 0;
            up_po_flag = 0;
            check_invoice_flag = 0;
            up_invoice_flag = 0;
            check_fcst_flag = 0;
            up_fcst_flag = 0;
            check_plan_flag = 0;
            up_plan_flag = 1;
            if (!backgroundWorker1.IsBusy)
            {
                backgroundWorker1.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Tiến trình khác đang chạy, thử lại sau");
            }
        }

        private void sửaPOToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            suaPO();
        }

        private void sửaInvoiceToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            suainvoice();
        }

        private void xóaPOToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            xoaPO();
        }

        private void xóaInvoiceToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            xoaInvoice();
        }

        private void xóaPlanToolStripMenuItem_Click(object sender, EventArgs e)
        {
            xoaplan();
        }

        private void xóaFCSTToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            xoaFCST();
        }

        private void thêmGiaoHàngChoPOToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (po_flag == 1)
            {
                InvoiceForm invoiceform = new InvoiceForm();
                invoiceform.loginIDInvoiceForm = LoginID;

                var selectedRows = dataGridView1.SelectedRows
                      .OfType<DataGridViewRow>()
                      .Where(row => !row.IsNewRow)
                      .ToArray();
                foreach (var row in selectedRows)
                {
                    try
                    {
                        string CUST_CD, EMPL_NO, G_CODE, PO_NO, DELIVERY_QTY, DELIVERY_DATE, NOCANCEL, REMARK, G_NAME, CUST_NAME;
                        string CTR_CD = "002";
                        G_CODE = row.Cells["G_CODE"].Value.ToString();
                        DELIVERY_QTY = "";
                        CUST_CD = row.Cells["CUST_CD"].Value.ToString();
                        EMPL_NO = row.Cells["EMPL_NO"].Value.ToString();
                        //DELIVERY_DATE = STYMD(DateTime.Today.Year, DateTime.Today.Month, DateTime.Today.Day); 
                        DateTime today = DateTime.Today;
                        today = today.AddDays(-1);
                        DELIVERY_DATE = today.ToString("yyyy-MM-dd");
                        PO_NO = row.Cells["PO_NO"].Value.ToString();
                        REMARK = "";
                        G_NAME = row.Cells["G_NAME"].Value.ToString();
                        CUST_NAME = row.Cells["CUST_NAME_KD"].Value.ToString();
                        NOCANCEL = "1";

                        invoiceform.CUST_CD = CUST_CD;
                        invoiceform.EMPL_NO = EMPL_NO;
                        invoiceform.G_CODE = G_CODE;
                        invoiceform.PO_NO = PO_NO;
                        invoiceform.DELIVERY_QTY = DELIVERY_QTY;
                        invoiceform.DELIVERY_DATE = DELIVERY_DATE;
                        invoiceform.NOCANCEL = NOCANCEL;
                        invoiceform.G_NAME = G_NAME;
                        invoiceform.CUST_NAME = CUST_NAME;


                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }


                }
                invoiceform.updateform();
                invoiceform.Show();
            }
            else
            {
                MessageBox.Show("Không phải bảng PO, ko thêm được giao hàng");
            }

        }

        private void tồnKiểmToolStripMenuItem_Click(object sender, EventArgs e)
        {
            checkCK();
        }

        private void tồnKiểmRútGọnToolStripMenuItem_Click(object sender, EventArgs e)
        {
            checkCK2();
        }

        private void bTPToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void bTPChiTiếtToolStripMenuItem_Click(object sender, EventArgs e)
        {
            checkBTP();
        }

        private void bTPRútGọnToolStripMenuItem_Click(object sender, EventArgs e)
        {
            checkBTP2();
        }

        private void thêmBOMAmazoneToolStripMenuItem_Click(object sender, EventArgs e)
        {
            BOMAMAZONE bOMAMAZONE = new BOMAMAZONE();
            bOMAMAZONE.Login_ID = LoginID;
            bOMAMAZONE.Show();
        }

        private void thiếtKếAmazoneToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("Đang phát triển, tạm thời dùng 1 design có sẵn");
            DESIGN_AMAZONE dESIGN_AMAZONE = new DESIGN_AMAZONE();
            dESIGN_AMAZONE.Login_ID = LoginID;
            dESIGN_AMAZONE.Show();
        }

        private void stockThànhPhẩmToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            KHOTHANHPHAM ktp = new KHOTHANHPHAM();
            ktp.Show();
        }

        public DataTable dtgv1_data = new DataTable();
        public DataTable dtgv1_total_data = new DataTable();

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            tracuu_KD(e);            
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {

        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            pictureBox1.Hide();
            if (po_flag == 1)
            {
                if (dtgv1_data.Rows.Count > 0)
                {
                    this.dataGridView1.DataSource = null;
                    this.dataGridView1.Columns.Clear();
                    dataGridView1.DataSource = dtgv1_data;
                    setRowNumber(dataGridView1);
                    button3.Enabled = true;                   

                    textBox2.Text = dtgv1_total_data.Rows[0]["PO_QTY"].ToString();
                    textBox2.Text = string.Format("{0:#,##0 EA}", double.Parse(textBox2.Text));
                    textBox4.Text = dtgv1_total_data.Rows[0]["TOTAL_DELIVERED"].ToString();
                    textBox4.Text = string.Format("{0:#,##0 EA}", double.Parse(textBox4.Text));
                    textBox6.Text = dtgv1_total_data.Rows[0]["PO_BALANCE"].ToString();
                    textBox6.Text = string.Format("{0:#,##0 EA}", double.Parse(textBox6.Text));

                    textBox3.Text = dtgv1_total_data.Rows[0]["PO_AMOUNT"].ToString();
                    textBox3.Text = string.Format("{0:#,##0.00 $}", double.Parse(textBox3.Text));
                    textBox5.Text = dtgv1_total_data.Rows[0]["DELIVERED_AMOUNT"].ToString();
                    textBox5.Text = string.Format("{0:#,##0.00 $}", double.Parse(textBox5.Text));
                    textBox7.Text = dtgv1_total_data.Rows[0]["BALANCE_AMOUNT"].ToString();
                    textBox7.Text = string.Format("{0:#,##0.00 $}", double.Parse(textBox7.Text));

                    textBox2.Font = new Font("Microsoft Sans Serif", 15.0f);
                    textBox4.Font = new Font("Microsoft Sans Serif", 15.0f);
                    textBox6.Font = new Font("Microsoft Sans Serif", 15.0f);
                    textBox3.Font = new Font("Microsoft Sans Serif", 15.0f);
                    textBox5.Font = new Font("Microsoft Sans Serif", 15.0f);
                    textBox7.Font = new Font("Microsoft Sans Serif", 15.0f);

                    formatDataGridViewtraPO1(dataGridView1);
                    /*
                    for (int r = 0; r < dataGridView1.Rows.Count; r++)
                    {
                        DataGridViewRow row = dataGridView1.Rows[r];
                        if (int.Parse(row.Cells["PO_BALANCE"].Value.ToString()) > 0 && (DateTime.Now - DateTime.Parse(row.Cells["PO_DATE"].Value.ToString())).Days >=90)
                        {
                           row.DefaultCellStyle.BackColor = Color.Red;
                            row.DefaultCellStyle.ForeColor = Color.White;
                        }
                    }
                    */
                    MessageBox.Show("Đã load  " + dtgv1_data.Rows.Count + "dòng");
                }
                else
                {
                    MessageBox.Show("Không có kết quả nào");
                }
                button8.Enabled = true;
            }
            else if (invoice_flag == 1)
            { 
                if (dtgv1_data.Rows.Count > 0)
                {
                    this.dataGridView1.DataSource = null;
                    this.dataGridView1.Columns.Clear();                   
                    dataGridView1.DataSource = dtgv1_data;
                    setRowNumber(dataGridView1);
                    button3.Enabled = true;
                  

                    textBox2.Text = "0";
                    textBox2.Text = string.Format("{0:#,##0 EA}", double.Parse(textBox2.Text));
                    textBox4.Text = dtgv1_total_data.Rows[0]["DELIVERED_QTY"].ToString();
                    textBox4.Text = string.Format("{0:#,##0 EA}", double.Parse(textBox4.Text));
                    textBox6.Text = "0";
                    textBox6.Text = string.Format("{0:#,##0 EA}", double.Parse(textBox6.Text));

                    textBox3.Text = "0";
                    textBox3.Text = string.Format("{0:#,##0.00 $}", double.Parse(textBox3.Text));
                    textBox5.Text = dtgv1_total_data.Rows[0]["DELIVERED_AMOUNT"].ToString();
                    textBox5.Text = string.Format("{0:#,##0.00 $}", double.Parse(textBox5.Text));
                    textBox7.Text = "0";
                    textBox7.Text = string.Format("{0:#,##0.00 $}", double.Parse(textBox7.Text));

                    textBox2.Font = new Font("Microsoft Sans Serif", 15.0f);
                    textBox4.Font = new Font("Microsoft Sans Serif", 15.0f);
                    textBox6.Font = new Font("Microsoft Sans Serif", 15.0f);
                    textBox3.Font = new Font("Microsoft Sans Serif", 15.0f);
                    textBox5.Font = new Font("Microsoft Sans Serif", 15.0f);
                    textBox7.Font = new Font("Microsoft Sans Serif", 15.0f);
                    formatDataGridViewtraInvoice1(dataGridView1);
                    MessageBox.Show("Đã load  " + dtgv1_data.Rows.Count + "dòng");                   
                }
                else
                {
                    MessageBox.Show("Không có kết quả nào");
                }

            }
            else if (fcst_flag == 1)
            {
                if (dtgv1_data.Rows.Count > 0)
                {
                    this.dataGridView1.DataSource = null;
                    this.dataGridView1.Columns.Clear();
                    dataGridView1.DataSource = dtgv1_data;
                    setRowNumber(dataGridView1);
                    formatDataGridViewtraFCST1(dataGridView1);
                    MessageBox.Show("Đã load  " + dtgv1_data.Rows.Count + "dòng");
                }
                else
                {
                    MessageBox.Show("Không có kết quả nào");
                }
            }
            else if (khgh_flag == 1)
            {               
                if (dtgv1_data.Rows.Count > 0)
                {
                    this.dataGridView1.DataSource = null;
                    this.dataGridView1.Columns.Clear();
                    dataGridView1.DataSource = dtgv1_data;
                    setRowNumber(dataGridView1);
                    formatDataGridViewtraPlan1(dataGridView1);
                    MessageBox.Show("Đã load  " + dtgv1_data.Rows.Count + "dòng");
                }
                else
                {
                    MessageBox.Show("Không có kết quả nào");
                }
            }
            else if (ycsx_flag == 1)
            {
                
            }
            else if (bom_flag == 1)
            {               

                if (dtgv1_data.Rows.Count > 0)
                {
                    this.dataGridView1.DataSource = null;
                    this.dataGridView1.Columns.Clear();
                    dataGridView1.DataSource = dtgv1_data;
                    MessageBox.Show("Đã load " + dtgv1_data.Rows.Count + " dòng, double click vào code để xem BOM");
                }
                else
                {
                    MessageBox.Show("Không có code này !");
                }

            }
        }

        private void button36_Click(object sender, EventArgs e)
        {
            dataGridView1.Columns.Clear();

            po_flag = 1;
            bom_flag = 0;
            fcst_flag = 0;
            invoice_flag = 0;
            ycsx_flag = 0;
            khgh_flag = 0;

            import_excel_flag = 0;
            check_po_flag = 0;
            up_po_flag = 0;
            check_invoice_flag = 0;
            up_invoice_flag = 0;
            check_fcst_flag = 0;
            up_fcst_flag = 0;
            check_plan_flag = 0;
            up_plan_flag = 0;
            if (!backgroundWorker1.IsBusy)
            {
                backgroundWorker1.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Tiến trình khác đang chạy, thử lại sau");
            }
        }

        private void button38_Click(object sender, EventArgs e)
        {
            dataGridView1.Columns.Clear();
            po_flag = 0;
            bom_flag = 0;
            fcst_flag = 0;
            invoice_flag = 1;
            ycsx_flag = 0;
            khgh_flag = 0;

            import_excel_flag = 0;
            check_po_flag = 0;
            up_po_flag = 0;
            check_invoice_flag = 0;
            up_invoice_flag = 0;
            check_fcst_flag = 0;
            up_fcst_flag = 0;
            check_plan_flag = 0;
            up_plan_flag = 0;
            if (!backgroundWorker1.IsBusy)
            {
                backgroundWorker1.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Tiến trình khác đang chạy, thử lại sau");
            }
        }

        private void button37_Click(object sender, EventArgs e)
        {
            dataGridView1.Columns.Clear();
            po_flag = 0;
            bom_flag = 0;
            fcst_flag = 1;
            invoice_flag = 0;
            ycsx_flag = 0;
            khgh_flag = 0;

            import_excel_flag = 0;
            check_po_flag = 0;
            up_po_flag = 0;
            check_invoice_flag = 0;
            up_invoice_flag = 0;
            check_fcst_flag = 0;
            up_fcst_flag = 0;
            check_plan_flag = 0;
            up_plan_flag = 0;
            if (!backgroundWorker1.IsBusy)
            {
                backgroundWorker1.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Tiến trình khác đang chạy, thử lại sau");
            }
        }

        private void button35_Click(object sender, EventArgs e)
        {
            dataGridView1.Columns.Clear();
            po_flag = 0;
            bom_flag = 0;
            fcst_flag = 0;
            invoice_flag = 0;
            ycsx_flag = 0;
            khgh_flag = 1;

            import_excel_flag = 0;
            check_po_flag = 0;
            up_po_flag = 0;
            check_invoice_flag = 0;
            up_invoice_flag = 0;
            check_fcst_flag = 0;
            up_fcst_flag = 0;
            check_plan_flag = 0;
            up_plan_flag = 0;
            if (!backgroundWorker1.IsBusy)
            {
                backgroundWorker1.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Tiến trình khác đang chạy, thử lại sau");
            }
        }

        private void button39_Click(object sender, EventArgs e)
        {
            dataGridView1.Columns.Clear();
            po_flag = 0;
            bom_flag = 1;
            fcst_flag = 0;
            invoice_flag = 0;
            ycsx_flag = 0;
            khgh_flag = 0;

            import_excel_flag = 0;
            check_po_flag = 0;
            up_po_flag = 0;
            check_invoice_flag = 0;
            up_invoice_flag = 0;
            check_fcst_flag = 0;
            up_fcst_flag = 0;
            check_plan_flag = 0;
            up_plan_flag = 0;

            if (!backgroundWorker1.IsBusy)
            {
                backgroundWorker1.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Tiến trình khác đang chạy, thử lại sau");
            }
        }

        private void checkCodeThiếuDataCapaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            po_flag = 0;
            invoice_flag = 0;
            khgh_flag = 0;
            fcst_flag = 0;
            ycsx_flag = 0;
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dt = pro.report_QLSX();
            dataGridView1.Columns.Clear();
            this.dataGridView1.DataSource = null;
            this.dataGridView1.Rows.Clear();
            this.dataGridView1.DataSource = dt;
            MessageBox.Show("Đã load : " + dt.Rows.Count + " dòng");
            MessageBox.Show("Đã load danh sách các code còn tồn PO, có xuất hiện trong kế hoạch giao hàng SOP 14 ngày trở lại đây, có xuất hiện trong phòng kiểm tra 14 ngày trở lại đây. Mà vẫn chưa được update thông tin");
        }

        private void checkCodeSaiQuyTắcCapaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            po_flag = 0;
            invoice_flag = 0;
            khgh_flag = 0;
            fcst_flag = 0;
            ycsx_flag = 0;
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dt = pro.report_QLSX_validating();
            dataGridView1.Columns.Clear();
            this.dataGridView1.DataSource = null;
            this.dataGridView1.Rows.Clear();
            this.dataGridView1.DataSource = dt;
            MessageBox.Show("Đã load : " + dt.Rows.Count + " dòng");
            MessageBox.Show("Đã load danh sách các code điền ko đúng quy tắc");
        }

        private void updateDatadòngĐượcChọnToolStripMenuItem_Click(object sender, EventArgs e)
        {
            upthongtincodeQLSXhangloat();
        }

        public void traBOMCAPA()
        {
            po_flag = 0;
            invoice_flag = 0;
            khgh_flag = 0;
            fcst_flag = 0;
            ycsx_flag = 0;
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dt = pro.traBOMCAPA(traBOMcondition());
            dataGridView1.Columns.Clear();
            this.dataGridView1.DataSource = null;
            this.dataGridView1.Rows.Clear();
            this.dataGridView1.DataSource = dt;
            MessageBox.Show("Đã load : " + dt.Rows.Count + " dòng");
           

        }
        public string traBOMcondition()
        {
            string condition = " WHERE 1=1";
            string G_CODE = "";
            if(textBox16.Text != "")
            {
                G_CODE = $" AND G_CODE= '{textBox16.Text}'";
            }
            string G_NAME = "";
            if (textBox1.Text != "")
            {
                G_CODE = $" AND G_NAME LIKE '%{textBox1.Text}%'";
            }
            condition += G_CODE + G_NAME;
            return condition;
        }

        private void traBOMCapaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            traBOMCAPA();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            ProductBLL pro = new ProductBLL();
            pro.updateOnline(LoginID);
        }

        private void dataGridView1_Sorted(object sender, EventArgs e)
        {
            /*
            if(po_flag == 1)
            {
                for (int r = 0; r < dataGridView1.Rows.Count; r++)
                {
                    DataGridViewRow row = dataGridView1.Rows[r];
                    if (int.Parse(row.Cells["PO_BALANCE"].Value.ToString()) > 0 && (DateTime.Now - DateTime.Parse(row.Cells["RD_DATE"].Value.ToString())).Days >= 60)
                    {
                        row.DefaultCellStyle.BackColor = Color.Red;
                        row.DefaultCellStyle.ForeColor = Color.White;
                    }
                }
            }
            */
            
        }

        private void checkToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LICHSUSANXUAT lssx = new LICHSUSANXUAT();
            lssx.Show();
        }

        private void kinhDoanhToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("Hãy mở web");
        }

        private void xóaYêuCầuSảnXuấtHàngLoạtToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            if (textBox8.Text == "xoa")
            {
                if (MessageBox.Show("Bạn thực sự muốn xóa YCSX?", "Xóa YCSX ?", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    try
                    {
                        ProductBLL pro = new ProductBLL();
                        DataTable dt = new DataTable();

                        var selectedRows = dataGridView1.SelectedRows
                       .OfType<DataGridViewRow>()
                       .Where(row => !row.IsNewRow)
                       .ToArray();

                        progressBar1.Minimum = 0; //Đặt giá trị nhỏ nhất cho ProgressBar
                        progressBar1.Maximum = dataGridView1.SelectedRows.Count; //Đặt giá trị lớn nhất cho ProgressBar
                        int startprogress = 0;

                        foreach (var row in selectedRows)
                        {
                            string ycsxno = row.Cells[3].Value.ToString();
                            dt = pro.checkYCSXO300(ycsxno);
                            if(dt.Rows.Count ==0)
                            {
                                dt = pro.DeleteYCSX(ycsxno);
                                pro.writeHistory("002", LoginID, "YCSX TABLE", "XOA", "XOA YCSX", "0");
                            }                            
                            startprogress = startprogress + 1;
                            label4.Text = "Progress: " + startprogress + "/" + dataGridView1.SelectedRows.Count;
                            progressBar1.Value = startprogress;
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
        }

        public void xoaPO()
        { 
            if(po_flag==1)
            {
                if (MessageBox.Show("Bạn thực sự muốn xóa PO?", "Xóa PO ?", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    try
                    {
                        ProductBLL pro = new ProductBLL();
                        DataTable dt = new DataTable();

                        var selectedRows = dataGridView1.SelectedRows
                       .OfType<DataGridViewRow>()
                       .Where(row => !row.IsNewRow)
                       .ToArray();

                        progressBar1.Minimum = 0; //Đặt giá trị nhỏ nhất cho ProgressBar
                        progressBar1.Maximum = dataGridView1.SelectedRows.Count; //Đặt giá trị lớn nhất cho ProgressBar
                        int startprogress = 0;

                        foreach (var row in selectedRows)
                        {
                            string po_id = row.Cells["PO_ID"].Value.ToString();//cell 22
                            try
                            {
                                dt = pro.DeletePO(po_id);
                                pro.writeHistory("002", LoginID, "PO TABLE", "XOA", "XOA PO", po_id);
                                //MessageBox.Show("Xóa PO thành công !");
                                startprogress = startprogress + 1;
                                label4.Text = "Progress: " + startprogress + "/" + dataGridView1.SelectedRows.Count;
                                progressBar1.Value = startprogress;
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Không thể xóa PO đã phát sinh giao hàng !");
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
                MessageBox.Show("Không phải bảng PO nên không xóa được !");
            }
                
           
        }
        private void xóaPOĐãChọnToolStripMenuItem_Click(object sender, EventArgs e)
        {
            xoaPO();
        }

        private void tạo1InvoiceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            InvoiceForm invoiceform = new InvoiceForm();
            invoiceform.loginIDInvoiceForm = LoginID;           
            invoiceform.Show();
        }

        private void dataGridView1_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            
            if(po_flag == 1)
            {
                InvoiceForm invoiceform = new InvoiceForm();
                invoiceform.loginIDInvoiceForm = LoginID;

                var selectedRows = dataGridView1.SelectedRows
                      .OfType<DataGridViewRow>()
                      .Where(row => !row.IsNewRow)
                      .ToArray();
                foreach (var row in selectedRows)
                {                  
                    try
                    {
                        string CUST_CD, EMPL_NO, G_CODE, PO_NO, DELIVERY_QTY, DELIVERY_DATE, NOCANCEL, REMARK, G_NAME, CUST_NAME;
                        string CTR_CD = "002";
                        G_CODE = row.Cells["G_CODE"].Value.ToString();
                        DELIVERY_QTY = "";
                        CUST_CD = row.Cells["CUST_CD"].Value.ToString(); 
                        EMPL_NO = row.Cells["EMPL_NO"].Value.ToString();
                        //DELIVERY_DATE = STYMD(DateTime.Today.Year, DateTime.Today.Month, DateTime.Today.Day); 
                        DateTime today = DateTime.Today;
                        today = today.AddDays(-1);
                        DELIVERY_DATE = today.ToString("yyyy-MM-dd");
                        PO_NO = row.Cells["PO_NO"].Value.ToString(); 
                        REMARK = "";
                        G_NAME = row.Cells["G_NAME"].Value.ToString();
                        CUST_NAME = row.Cells["CUST_NAME"].Value.ToString();
                        NOCANCEL = "1";

                        invoiceform.CUST_CD = CUST_CD;
                        invoiceform.EMPL_NO = EMPL_NO;
                        invoiceform.G_CODE = G_CODE;
                        invoiceform.PO_NO = PO_NO;
                        invoiceform.DELIVERY_QTY = DELIVERY_QTY;
                        invoiceform.DELIVERY_DATE = DELIVERY_DATE;
                        invoiceform.NOCANCEL = NOCANCEL;
                        invoiceform.G_NAME = G_NAME;
                        invoiceform.CUST_NAME = CUST_NAME;


                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }


                }
                invoiceform.updateform();
                invoiceform.Show();
                
            }
            else if(bom_flag==1)
            {
                bom_flag = 0;
                var selectedRows = dataGridView1.SelectedRows
                     .OfType<DataGridViewRow>()
                     .Where(row => !row.IsNewRow)
                     .ToArray();
                foreach (var row in selectedRows)
                {
                    try
                    {
                        string G_CODE = row.Cells[0].Value.ToString();
                        ProductBLL pro = new ProductBLL();
                        DataTable dt = new DataTable();
                        dt = pro.info_getCODEBOM(G_CODE);
                        if(dt.Rows.Count>0)
                        {
                            dataGridView1.DataSource = dt;
                        }
                        else
                        {
                            MessageBox.Show("Khong co BOM");
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


    public class loginIDName
    {
        public string loginID;
    }
        
    
    public class YeuCauSanXuat
    {
        public string PROD_REQUEST_DATE { get; set; }
        public string CODE_50 { get; set; }
        public string CODE_55 { get; set; }
        public string G_CODE { get; set; }
        public string RIV_NO { get; set; }
        public string PROD_REQUEST_QTY { get; set; }
        public string CUST_CD { get; set; }
        public string EMPL_NO { get; set; }
        public string REMK { get; set; }
        public string DELIVERY_DT { get; set; }

        public string ext1 { get; set; }
        public string ext2 { get; set; }
        public string ext3 { get; set; }
        public string ext4 { get; set; }
        public string ext5 { get; set; }
        public string ext6 { get; set; }
        public string ext7 { get; set; }
        public string ext8 { get; set; }
        public string ext9 { get; set; }
        public string ext10 { get; set; }
        public string ext11 { get; set; }
        public string ext12 { get; set; }
        public string ext13 { get; set; }
        public string ext14 { get; set; }
        public string ext15 { get; set; }
        public string ext16 { get; set; }
        public string ext17 { get; set; }
        public string ext18 { get; set; }
        public string ext19 { get; set; }
        public string ext20 { get; set; }

    }
    public class codefullInfo
    {
        public string PROD_REQUEST_DATE { get; set; }
        public string CODE_50 { get; set; }
        public string CODE_55 { get; set; }
        public string G_CODE { get; set; }
        public string RIV_NO { get; set; }
        public string PROD_REQUEST_QTY { get; set; }
        public string CUST_CD { get; set; }
        public string EMPL_NO { get; set; }
        public string REMK { get; set; }
        public string DELIVERY_DT { get; set; }
    }



    public class ExcelFactory
    {

        public static void fastexport(DataTable dt,string path)
        {            /*Set up work book, work sheets, and excel application*/
            Microsoft.Office.Interop.Excel.Application oexcel = new Microsoft.Office.Interop.Excel.Application();
            try
            {                
                eX.Workbook MyBook = null;
                Microsoft.Office.Interop.Excel.Application MyApp = null;
                eX.Worksheet MySheet = null;
                MyApp = new Microsoft.Office.Interop.Excel.Application();
                MyApp.Visible = false;
                MyBook = MyApp.Workbooks.Open(path);
                MySheet = (eX.Worksheet)MyBook.Sheets["Report1"];


                //  obook.Worksheets.Add(misValue);

                
                int colIndex = 0;
                int rowIndex = 1;

                foreach (DataColumn dc in dt.Columns)
                {
                    colIndex++;
                    MySheet.Cells[1, colIndex] = dc.ColumnName;
                }
                foreach (DataRow dr in dt.Rows)
                {
                    rowIndex++;
                    colIndex = 0;

                    foreach (DataColumn dc in dt.Columns)
                    {
                        colIndex++;
                        MySheet.Cells[rowIndex, colIndex] = dr[dc.ColumnName];
                    }
                }

                MySheet.Columns.AutoFit();                

                //Release and terminate excel
                string Dir = System.IO.Directory.GetCurrentDirectory();
                DateTime hientai = DateTime.Today;
                string homnay = hientai.ToString();
                string file = Dir + "\\REPORT.xlsx";

                MySheet.SaveAs(file);

               

                MyBook.Close();
                oexcel.Quit();
                releaseObject(MySheet);

                releaseObject(MyBook);

                releaseObject(oexcel);
                GC.Collect();
            }
            catch (Exception ex)
            {
                oexcel.Quit();               
            }
        }


        public static void editFileExcelReport(string path, DataTable dt, string saveycsxpath)
        {
            try
            {
                eX.Workbook MyBook = null;
                Microsoft.Office.Interop.Excel.Application MyApp = null;
                eX.Worksheet MySheet = null;
                MyApp = new Microsoft.Office.Interop.Excel.Application();
                MyApp.Visible = false;
                MyBook = MyApp.Workbooks.Open(path);
                MySheet = (eX.Worksheet)MyBook.Sheets["Report1"];

                int numRow = dt.Rows.Count;
                int numCol = dt.Columns.Count;

                int colIndex = 0;
                int rowIndex = 1;

                foreach (DataColumn dc in dt.Columns)
                {
                    colIndex++;
                    MySheet.Cells[1, colIndex] = dc.ColumnName;
                }
                foreach (DataRow dr in dt.Rows)
                {
                    rowIndex++;
                    colIndex = 0;

                    foreach (DataColumn dc in dt.Columns)
                    {
                        colIndex++;
                        MySheet.Cells[rowIndex, colIndex] = dr[dc.ColumnName];
                    }
                }



                /*
                MySheet.Range["A64"].Value = dt.Rows[0]["REMK"].ToString();
                MySheet.Range["AA1"].Value = dt.Rows[0]["PROD_REQUEST_DATE"].ToString().Substring(2) + "-" + dt.Rows[0]["PROD_REQUEST_NO"].ToString();
                MySheet.Range["AA3"].Value = "*" + dt.Rows[0]["PROD_REQUEST_NO"].ToString() + "*";
                MySheet.Range["F4"].Value = dt.Rows[0]["EMPL_NAME"].ToString();
                MySheet.Range["F5"].Value = dt.Rows[0]["CUST_NAME"].ToString();
                MySheet.Range["F6"].Value = dt.Rows[0]["G_CODE"].ToString();
                MySheet.Range["F7"].Value = dt.Rows[0]["G_NAME"].ToString();
                MySheet.Range["I6"].Value = "*" + dt.Rows[0]["G_CODE"].ToString() + "*";

                MySheet.Range["T4"].Value = dt.Rows[0]["PROD_REQUEST_QTY"].ToString();
                MySheet.Range["T6"].Value = dt.Rows[0]["PROD_REQUEST_QTY"].ToString();
                MySheet.Range["T7"].Value = dt.Rows[0]["DELIVERY_DT"].ToString();
                */

                object misValue = System.Reflection.Missing.Value;
                MyBook.Close(true, saveycsxpath, misValue);
                MyApp.Quit();
                releaseObject(MySheet);
                releaseObject(MyBook);
                releaseObject(MyApp);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }




        public static void printPDF(string pdffilepath)
        {
            ProcessStartInfo info = new ProcessStartInfo();
            info.Verb = "print";
            info.FileName = pdffilepath;
            info.CreateNoWindow = true;
            info.WindowStyle = ProcessWindowStyle.Normal;
            Process p = new Process();
            p.StartInfo = info;
            p.Start();
            if (!p.WaitForExit(500))
            {
                p.Kill();
            }
        }




        public static void editFileExcel(string path, DataTable dt, CheckBox ckb, string saveycsxpath)
        {
            try
            {
                eX.Workbook MyBook = null;
                Microsoft.Office.Interop.Excel.Application MyApp = null;
                eX.Worksheet MySheet = null;
                MyApp = new Microsoft.Office.Interop.Excel.Application();
                MyApp.Visible = false;
                MyBook = MyApp.Workbooks.Open(path);
                MySheet = (eX.Worksheet)MyBook.Sheets[1];

                int numRow = dt.Rows.Count;
                int numCol = dt.Columns.Count;

                MySheet.Range["A64"].Value=dt.Rows[0]["REMK"].ToString();
                MySheet.Range["AA1"].Value =   dt.Rows[0]["PROD_REQUEST_DATE"].ToString().Substring(2) + "-" +dt.Rows[0]["PROD_REQUEST_NO"].ToString();
                MySheet.Range["AA3"].Value = "*" + dt.Rows[0]["PROD_REQUEST_NO"].ToString() + "*";
                MySheet.Range["F4"].Value = dt.Rows[0]["EMPL_NAME"].ToString();
                MySheet.Range["F5"].Value = dt.Rows[0]["CUST_NAME"].ToString();
                MySheet.Range["F6"].Value = dt.Rows[0]["G_CODE"].ToString();
                MySheet.Range["F7"].Value = dt.Rows[0]["G_NAME"].ToString();
                MySheet.Range["I6"].Value = "*"+dt.Rows[0]["G_CODE"].ToString()+"*";

                MySheet.Range["T4"].Value = dt.Rows[0]["PROD_REQUEST_QTY"].ToString();                
                MySheet.Range["T6"].Value = dt.Rows[0]["PROD_REQUEST_QTY"].ToString();
                MySheet.Range["T7"].Value = dt.Rows[0]["DELIVERY_DT"].ToString();

                //Phan loai san xuat
                switch(dt.Rows[0]["CODE_55"].ToString())
                {
                    case "01": //thong thuong
                        MySheet.Range["AE5"].Value = "Thong Thuong";
                        break;
                    case "02": //SDI.
                        MySheet.Range["AE5"].Value = "SDI";
                        break;
                    case "03": //GC
                        MySheet.Range["AE5"].Value = "GC";  
                        break;
                    case "04": //SAMPLE
                        MySheet.Range["AE5"].Value = "SAMPLE";
                        break;
                    default:
                        break;  
                }

                // Phan loai giao hang
                switch (dt.Rows[0]["CODE_50"].ToString())
                {
                    case "01": //GC
                        MySheet.Range["AE7"].Value = "Thong Thuong";
                        break;
                    case "02": //SK
                        MySheet.Range["AE7"].Value = "SK";
                        break;
                    case "03": //KD
                        MySheet.Range["AE7"].Value = "KD";
                        break;
                    case "04": //VN
                        MySheet.Range["AE7"].Value = "VN";
                        break;
                    case "05": //SAMPLE
                        MySheet.Range["AE7"].Value = "SAMPLE";
                        break;
                    case "06": //Vai bac 4
                        MySheet.Range["AE7"].Value = "Vai bac 4";
                        break;
                    case "07": //ETC
                        MySheet.Range["AE7"].Value = "ETC";
                        break;
                    default:
                        break;
                }

                MySheet.Range["E10"].Value = dt.Rows[0]["G_WIDTH"].ToString()+"mm";
                MySheet.Range["E11"].Value = dt.Rows[0]["G_LENGTH"].ToString() + "mm";
                MySheet.Range["E12"].Value = dt.Rows[0]["G_R"].ToString();
                MySheet.Range["E13"].Value = dt.Rows[0]["G_C"].ToString() + "EA";

                //packing type
                switch (dt.Rows[0]["CODE_33"].ToString())   
                {
                    case "01": //EA
                        MySheet.Range["Q10"].Value =  "EA";
                        break;
                    case "02": //Roll
                        MySheet.Range["Q10"].Value = "ROLL";
                        break;
                    case "03": //Sheet
                        MySheet.Range["Q10"].Value = "SHEET";
                        break;
                    case "04": //Met
                        MySheet.Range["Q10"].Value = "MET";
                        break;
                    case "06": //Pack
                        MySheet.Range["Q10"].Value = "PACK (BAG)";
                        break;
                    case "99": //X
                        MySheet.Range["Q10"].Value = "X";
                        break;                    
                    default:
                        break;
                }

                //MySheet.Range["Q10"].Value = dt.Rows[0]["CODE_33"].ToString();// packing type
                MySheet.Range["Q11"].Value = dt.Rows[0]["ROLE_EA_QTY"].ToString()+"EA"; // packing qty               
                MySheet.Range["Q13"].Value = dt.Rows[0]["PACK_DRT"].ToString();

                MySheet.Range["AB10"].Value = dt.Rows[0]["G_LG"].ToString() + "mm";// packing type
                MySheet.Range["AB11"].Value = dt.Rows[0]["G_SG_L"].ToString() + "mm"; // packing qty
                MySheet.Range["AB12"].Value = dt.Rows[0]["G_SG_R"].ToString() + "mm";

                double req_qty = double.Parse(dt.Rows[0]["PROD_REQUEST_QTY"].ToString());
                double dai = double.Parse(dt.Rows[0]["G_LENGTH"].ToString());
                double gap = double.Parse(dt.Rows[0]["G_LG"].ToString());
                double cavity = double.Parse(dt.Rows[0]["G_C"].ToString());
                double metdai = (dai + gap) * req_qty / cavity / 1000;

                MySheet.Range["AB15"].Value = metdai.ToString()+"m";
                int startRow = 17;

                //MessageBox.Show("Number of Rows =" + numRow);
                for(int i= 0;i< numRow;i++)
                {                    
                    int afterstartrow0 = startRow;
                    int afterstartrow1 = startRow + 1;
                    //MessageBox.Show("F" + afterstartrow0);
                    MySheet.Range["F" + afterstartrow0].Value = dt.Rows[i]["M_CODE"].ToString(); //material code
                    //MessageBox.Show(MySheet.Range["F" + startRow + i].Value);
                    MySheet.Range["F" + afterstartrow1].Value = "*"+dt.Rows[i]["M_CODE"].ToString() + "*";
                    //MessageBox.Show(MySheet.Range["F" + startRow + i + 1].Value);
                    MySheet.Range["M" + afterstartrow0].Value = dt.Rows[i]["M_NAME"].ToString(); //material name
                   // MessageBox.Show(MySheet.Range["M" + startRow + i].Value);
                    MySheet.Range["Z" + afterstartrow0].Value = dt.Rows[i]["WIDTH_CD"].ToString();
                    // MessageBox.Show(MySheet.Range["Z" + startRow + i].Value);
                    MySheet.Range["AJ" + afterstartrow0].Value = dt.Rows[i]["REMARK"].ToString();
                    startRow += 2;

                }
                startRow = 17;

                MySheet.Range["AA76"].Value = MySheet.Range["AB10"].Value;// label gap
                MySheet.Range["X81"].Value = MySheet.Range["E10"].Value; // width
                MySheet.Range["V83"].Value = MySheet.Range["E11"].Value; //length
                MySheet.Range["V86"].Value = MySheet.Range["AB12"].Value; //left side gap
                MySheet.Range["AF86"].Value = MySheet.Range["AB11"].Value; //right side gap

                object misValue = System.Reflection.Missing.Value;
                string path2 = saveycsxpath + "\\" + dt.Rows[0]["PROD_REQUEST_NO"].ToString() + "-" + dt.Rows[0]["G_NAME"].ToString().Substring(0,11) + ".xlsx";
                //string path2 = "C:\\Users\\PQC_NM2\\Desktop\\ycsx\\" + dt.Rows[0]["PROD_REQUEST_NO"].ToString() + "-" + dt.Rows[0]["G_NAME"].ToString() + ".xlsx";
                //MessageBox.Show(ckb.Checked.ToString());             
                if(ckb.Checked.ToString()=="True")
                {
                    MySheet.PrintOut(Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    string Dir = System.IO.Directory.GetCurrentDirectory();
                    string drawfilename = dt.Rows[0]["G_NAME"].ToString().Substring(0, 11)+"_" + dt.Rows[0]["G_CODE"].ToString().Substring(7,1)+".pdf";
                    //MessageBox.Show(drawfilename);
                    string file = Dir + "\\BANVE\\"+drawfilename;
                    //MessageBox.Show(file);
                    if (File.Exists(file))
                    {
                       // printPDF(file);
                    }
                    else
                    {
                        MessageBox.Show("Không có bản vẽ : " + dt.Rows[0]["G_NAME"].ToString());
                    }
                }

                MyBook.Close(true, path2, misValue);
                MyApp.Quit();
                releaseObject(MySheet);
                releaseObject(MyBook);
                releaseObject(MyApp);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void editFileBOMExcel(string path, DataTable dt, string saveycsxpath, string draw_link)
        {
            try
            {
                eX.Workbook MyBook = null;
                Microsoft.Office.Interop.Excel.Application MyApp = null;
                eX.Worksheet MySheet = null;
                MyApp = new Microsoft.Office.Interop.Excel.Application();
                MyApp.Visible = false;
                MyBook = MyApp.Workbooks.Open(path);
                MySheet = (eX.Worksheet)MyBook.Sheets[1];

                int numRow = dt.Rows.Count;
                int numCol = dt.Columns.Count;

                MySheet.Range["K5"].Value = dt.Rows[0]["M100_INS_EMPL"].ToString();
                MySheet.Range["K6"].Value = dt.Rows[0]["M110_CUST_NAME_KD"].ToString();
                MySheet.Range["K7"].Value = dt.Rows[0]["M100_PROD_PROJECT"].ToString();
                MySheet.Range["AD5"].Value = dt.Rows[0]["M100_PROD_MODEL"].ToString();
                MySheet.Range["AD6"].Value = dt.Rows[0]["M100_PROD_TYPE"].ToString();
                MySheet.Range["AD7"].Value = dt.Rows[0]["M100_G_NAME_KD"].ToString();
                MySheet.Range["AS5"].Value = dt.Rows[0]["M100_DESCR"].ToString();
                MySheet.Range["AS6"].Value = dt.Rows[0]["M100_PROD_MAIN_MATERIAL"].ToString();
                MySheet.Range["AS7"].Value = dt.Rows[0]["M100_G_CODE"].ToString();

                MySheet.Range["G11"].Value = dt.Rows[0]["M100_G_WIDTH"].ToString();
                MySheet.Range["O11"].Value = dt.Rows[0]["M100_G_LENGTH"].ToString();
                MySheet.Range["K12"].Value = dt.Rows[0]["M100_PD"].ToString();
                MySheet.Range["G13"].Value = dt.Rows[0]["M100_G_C"].ToString();
                MySheet.Range["O13"].Value = dt.Rows[0]["M100_G_C_R"].ToString();
                MySheet.Range["G14"].Value = dt.Rows[0]["M100_G_CG"].ToString();
                MySheet.Range["O14"].Value = dt.Rows[0]["M100_G_LG"].ToString();

                MySheet.Range["Z11"].Value = dt.Rows[0]["M100_G_SG_L"].ToString();
                MySheet.Range["AH11"].Value = dt.Rows[0]["M100_G_SG_R"].ToString();
                MySheet.Range["AD12"].Value = (dt.Rows[0]["M100_KNIFE_TYPE"].ToString()== "0" ? "PVC" : dt.Rows[0]["M100_KNIFE_TYPE"].ToString()=="1" ? "PINACLE" : "NO");
                MySheet.Range["AD13"].Value = dt.Rows[0]["M100_KNIFE_LIFECYCLE"].ToString();       
                switch (dt.Rows[0]["M100_CODE_33"].ToString())
                {
                    case "01": //EA
                        MySheet.Range["AD14"].Value = "EA";
                        break;
                    case "02": //Roll
                        MySheet.Range["AD14"].Value = "ROLL";
                        break;
                    case "03": //Sheet
                        MySheet.Range["AD14"].Value = "SHEET";
                        break;
                    case "04": //Met
                        MySheet.Range["AD14"].Value = "MET";
                        break;
                    case "06": //Pack
                        MySheet.Range["AD14"].Value = "PACK (BAG)";
                        break;
                    case "99": //X
                        MySheet.Range["AD14"].Value = "X";
                        break;
                    default:
                        break;
                }

                MySheet.Range["AW11"].Value = dt.Rows[0]["M100_RPM"].ToString();
                MySheet.Range["AW12"].Value = dt.Rows[0]["M100_PACK_DRT"].ToString();
                MySheet.Range["AW13"].Value = dt.Rows[0]["M100_PIN_DISTANCE"].ToString();
                MySheet.Range["AW14"].Value = dt.Rows[0]["M100_PROCESS_TYPE"].ToString();

                MySheet.Range["K18"].Value = dt.Rows[0]["M100_EQ1"].ToString();
                MySheet.Range["K19"].Value = dt.Rows[0]["M100_EQ2"].ToString();

                MySheet.Range["AD18"].Value = dt.Rows[0]["M100_PROD_DIECUT_STEP"].ToString();
                MySheet.Range["AD19"].Value = dt.Rows[0]["M100_PROD_PRINT_TIMES"].ToString();

                MySheet.Range["AW18"].Value = dt.Rows[0]["M100_UPH1"].ToString();
                MySheet.Range["AW19"].Value = dt.Rows[0]["M100_UPH2"].ToString();

                MySheet.Range["K22"].Value = dt.Rows[0]["M100_G_CODE"].ToString();
                MySheet.Range["AD22"].Value = dt.Rows[0]["ZTB_BOM2_RIV_NO"].ToString();

                MySheet.Range["C37"].Value = "LINK: " + dt.Rows[0]["M100_DRAW_LINK"].ToString();

                MySheet.Range["AG2"].Value = "Created Date: " + DateTime.Today.ToString().Substring(0,10);



                int startRow = 26;

                //MessageBox.Show("Number of Rows =" + numRow);
                for (int i = 0; i < numRow; i++)
                {    
                    MySheet.Range["E" + (startRow+i)].Value = dt.Rows[i]["ZTB_BOM2_M_NAME"].ToString(); //material code
                    MySheet.Range["R" + (startRow + i)].Value = dt.Rows[i]["ZTB_BOM2_CUST_CD"].ToString(); //material code
                    MySheet.Range["X" + (startRow + i)].Value = dt.Rows[i]["ZTB_BOM2_USAGE"].ToString(); //material code
                    MySheet.Range["AA" + (startRow + i)].Value = dt.Rows[i]["ZTB_BOM2_MAT_MASTER_WIDTH"].ToString(); //material code
                    MySheet.Range["AD" + (startRow + i)].Value = dt.Rows[i]["ZTB_BOM2_MAT_CUTWIDTH"].ToString(); //material code
                    MySheet.Range["AG" + (startRow + i)].Value = dt.Rows[i]["ZTB_BOM2_MAT_ROLL_LENGTH"].ToString(); //material code
                    MySheet.Range["AJ" + (startRow + i)].Value = dt.Rows[i]["ZTB_BOM2_MAT_THICKNESS"].ToString(); //material code
                    MySheet.Range["AM" + (startRow + i)].Value = dt.Rows[i]["ZTB_BOM2_PROCESS_ORDER"].ToString(); //material code
                    MySheet.Range["AP" + (startRow + i)].Value = dt.Rows[i]["ZTB_BOM2_CATEGORY"].ToString(); //material code
                    MySheet.Range["AU" + (startRow + i)].Value = dt.Rows[i]["ZTB_BOM2_M_QTY"].ToString(); //material code
                }  

                object misValue = System.Reflection.Missing.Value;
                string path2 = saveycsxpath + "\\" + dt.Rows[0]["M100_G_CODE"].ToString()  + "_"+ dt.Rows[0]["M100_G_NAME"].ToString() + ".xlsx";

                //MySheet.PrintOut(Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                try
                {
                    MySheet.Shapes.AddOLEObject(Filename: $"{draw_link}", DisplayAsIcon: true, IconLabel: "DRAW_FILE", IconFileName: "", IconIndex: 0, Top: 580, Left: 20);

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi hệ điều hành, sẽ xuất BOM mà ko đính kèm bản vẽ");
                }
               
                //MySheet.Shapes.AddOLEObject(Filename: $"{draw_link}", Height: 10, Width: 10, Top: 580, Left: 20);

                MyBook.Close(true, path2, misValue);
                MyApp.Quit();
                releaseObject(MySheet);
                releaseObject(MyBook);
                releaseObject(MyApp);
            }            
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void editFileExcelQLSX(string path, DataTable dt, CheckBox ckb, string saveycsxpath)
        {
            try
            {
                eX.Workbook MyBook = null;
                Microsoft.Office.Interop.Excel.Application MyApp = null;
                eX.Worksheet MySheet = null;
                MyApp = new Microsoft.Office.Interop.Excel.Application();
                MyApp.Visible = false;
                MyBook = MyApp.Workbooks.Open(path);
                //MySheet = (eX.Worksheet)MyBook.Sheets[1];
                MySheet = (eX.Worksheet)MyBook.Worksheets["Info"];

                int numRow = dt.Rows.Count;
                int numCol = dt.Columns.Count;

                MySheet.Range["A64"].Value = dt.Rows[0]["REMK"].ToString();
                MySheet.Range["AA1"].Value = dt.Rows[0]["PROD_REQUEST_DATE"].ToString().Substring(2) + "-" + dt.Rows[0]["PROD_REQUEST_NO"].ToString();
                MySheet.Range["AA3"].Value = "*" + dt.Rows[0]["PROD_REQUEST_NO"].ToString() + "*";
                MySheet.Range["F4"].Value = dt.Rows[0]["EMPL_NAME"].ToString();
                MySheet.Range["F5"].Value = dt.Rows[0]["CUST_NAME"].ToString();
                MySheet.Range["F6"].Value = dt.Rows[0]["G_CODE"].ToString();
                MySheet.Range["F7"].Value = dt.Rows[0]["G_NAME"].ToString();
                MySheet.Range["I6"].Value = "*" + dt.Rows[0]["G_CODE"].ToString() + "*";

                MySheet.Range["T4"].Value = dt.Rows[0]["PROD_REQUEST_QTY"].ToString();
                MySheet.Range["T6"].Value = dt.Rows[0]["PROD_REQUEST_QTY"].ToString();
                MySheet.Range["T7"].Value = dt.Rows[0]["DELIVERY_DT"].ToString();

                //Phan loai san xuat
                switch (dt.Rows[0]["CODE_55"].ToString())
                {
                    case "01": //thong thuong
                        MySheet.Range["AE5"].Value = "Thong Thuong";
                        break;
                    case "02": //SDI.
                        MySheet.Range["AE5"].Value = "SDI";
                        break;
                    case "03": //GC
                        MySheet.Range["AE5"].Value = "GC";
                        break;
                    case "04": //SAMPLE
                        MySheet.Range["AE5"].Value = "SAMPLE";
                        break;
                    default:
                        break;
                }

                // Phan loai giao hang
                switch (dt.Rows[0]["CODE_50"].ToString())
                {
                    case "01": //GC
                        MySheet.Range["AE7"].Value = "Thong Thuong";
                        break;
                    case "02": //SK
                        MySheet.Range["AE7"].Value = "SK";
                        break;
                    case "03": //KD
                        MySheet.Range["AE7"].Value = "KD";
                        break;
                    case "04": //VN
                        MySheet.Range["AE7"].Value = "VN";
                        break;
                    case "05": //SAMPLE
                        MySheet.Range["AE7"].Value = "SAMPLE";
                        break;
                    case "06": //Vai bac 4
                        MySheet.Range["AE7"].Value = "Vai bac 4";
                        break;
                    case "07": //ETC
                        MySheet.Range["AE7"].Value = "ETC";
                        break;
                    default:
                        break;
                }

                MySheet.Range["E10"].Value = dt.Rows[0]["G_WIDTH"].ToString() + "mm";
                MySheet.Range["E11"].Value = dt.Rows[0]["G_LENGTH"].ToString() + "mm";
                MySheet.Range["E12"].Value = dt.Rows[0]["G_R"].ToString();
                MySheet.Range["E13"].Value = dt.Rows[0]["G_C"].ToString() + "EA";


                switch (dt.Rows[0]["CODE_33"].ToString())
                {
                    case "01": //EA
                        MySheet.Range["Q10"].Value = "EA";
                        break;
                    case "02": //Roll
                        MySheet.Range["Q10"].Value = "ROLL";
                        break;
                    case "03": //Sheet
                        MySheet.Range["Q10"].Value = "SHEET";
                        break;
                    case "04": //Met
                        MySheet.Range["Q10"].Value = "MET";
                        break;
                    case "06": //Pack
                        MySheet.Range["Q10"].Value = "PACK (BAG)";
                        break;
                    case "99": //X
                        MySheet.Range["Q10"].Value = "X";
                        break;
                    default:
                        break;
                }

                //MySheet.Range["Q10"].Value = dt.Rows[0]["CODE_33"].ToString();// packing type
                MySheet.Range["Q11"].Value = dt.Rows[0]["ROLE_EA_QTY"].ToString() + "EA"; // packing qty               
                MySheet.Range["Q13"].Value = dt.Rows[0]["PACK_DRT"].ToString();

                MySheet.Range["AB10"].Value = dt.Rows[0]["G_LG"].ToString() + "mm";// packing type
                MySheet.Range["AB11"].Value = dt.Rows[0]["G_SG_L"].ToString() + "mm"; // packing qty
                MySheet.Range["AB12"].Value = dt.Rows[0]["G_SG_R"].ToString() + "mm";

                double req_qty = double.Parse(dt.Rows[0]["PROD_REQUEST_QTY"].ToString());
                double dai = double.Parse(dt.Rows[0]["G_LENGTH"].ToString());
                double gap = double.Parse(dt.Rows[0]["G_LG"].ToString());
                double cavity = double.Parse(dt.Rows[0]["G_C"].ToString());
                double metdai = (dai + gap) * req_qty / cavity / 1000;

                MySheet.Range["AB15"].Value = metdai.ToString() + "m";
                int startRow = 17;

                //MessageBox.Show("Number of Rows =" + numRow);
                for (int i = 0; i < numRow; i++)
                {

                    int afterstartrow0 = startRow;
                    int afterstartrow1 = startRow + 1;
                    //MessageBox.Show("F" + afterstartrow0);
                    MySheet.Range["F" + afterstartrow0].Value = dt.Rows[i]["M_CODE"].ToString(); //material code
                    //MessageBox.Show(MySheet.Range["F" + startRow + i].Value);
                    MySheet.Range["F" + afterstartrow1].Value = "*" + dt.Rows[i]["M_CODE"].ToString() + "*";
                    //MessageBox.Show(MySheet.Range["F" + startRow + i + 1].Value);
                    MySheet.Range["M" + afterstartrow0].Value = dt.Rows[i]["M_NAME"].ToString(); //material name
                                                                                                 // MessageBox.Show(MySheet.Range["M" + startRow + i].Value);
                    MySheet.Range["Z" + afterstartrow0].Value = dt.Rows[i]["WIDTH_CD"].ToString();
                    // MessageBox.Show(MySheet.Range["Z" + startRow + i].Value);
                    MySheet.Range["AJ" + afterstartrow0].Value = dt.Rows[i]["REMARK"].ToString();
                    startRow += 2;

                }
                startRow = 17;

                MySheet.Range["AA76"].Value = MySheet.Range["AB10"].Value;// label gap
                MySheet.Range["X81"].Value = MySheet.Range["E10"].Value; // width
                MySheet.Range["V83"].Value = MySheet.Range["E11"].Value; //length
                MySheet.Range["V86"].Value = MySheet.Range["AB12"].Value; //left side gap
                MySheet.Range["AF86"].Value = MySheet.Range["AB11"].Value; //right side gap

                object misValue = System.Reflection.Missing.Value;
                string path2 = saveycsxpath + "\\" + dt.Rows[0]["PROD_REQUEST_NO"].ToString() + ".xlsx";
                //string path2 = "C:\\Users\\PQC_NM2\\Desktop\\ycsx\\" + dt.Rows[0]["PROD_REQUEST_NO"].ToString() + "-" + dt.Rows[0]["G_NAME"].ToString() + ".xlsx";


                if (ckb.Checked.ToString() == "True")
                {
                    MySheet.PrintOut(Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                }

                MyBook.Close(true, path2, misValue);
                MyApp.Quit();
                releaseObject(MySheet);
                releaseObject(MyBook);
                releaseObject(MyApp);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static DataTable readFromExcelFileToAmazoneTable(string path)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("DATA", typeof(string));
            try
            {
                eX.Workbook MyBook = null;
                Microsoft.Office.Interop.Excel.Application MyApp = null;
                eX.Worksheet MySheet = null;
                MyApp = new Microsoft.Office.Interop.Excel.Application();
                MyApp.Visible = false;
                MyBook = MyApp.Workbooks.Open(path);
                MySheet = (eX.Worksheet)MyBook.Sheets[1];

                int lastRow = MySheet.Cells.SpecialCells(eX.XlCellType.xlCellTypeLastCell).Row;
                int lastColumn = MySheet.Cells.SpecialCells(eX.XlCellType.xlCellTypeLastCell).Column;

                for (int i = 1; i <= lastRow; i++)
                {
                    YeuCauSanXuat nv = new YeuCauSanXuat();
                    string AMAZONE_DATA;
                    AMAZONE_DATA = MySheet.Range["A" + i].Value + "";
                    DataRow dr = dt.NewRow();
                    dr["DATA"] = AMAZONE_DATA;
                    dt.Rows.Add(dr);                                      
                }

                MyBook.Close(true);
                MyApp.Quit();
                releaseObject(MySheet);
                releaseObject(MyBook);
                releaseObject(MyApp);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return dt;
        }


        public static List<YeuCauSanXuat> readFromExcelFile(string path)
        {
            List<YeuCauSanXuat> dsNV = new List<YeuCauSanXuat>();
            try
            {
                eX.Workbook MyBook = null;
                Microsoft.Office.Interop.Excel.Application MyApp = null;
                eX.Worksheet MySheet = null;
                MyApp = new Microsoft.Office.Interop.Excel.Application();
                MyApp.Visible = false;
                MyBook = MyApp.Workbooks.Open(path);
                MySheet = (eX.Worksheet)MyBook.Sheets[1];

                int lastRow = MySheet.Cells.SpecialCells(eX.XlCellType.xlCellTypeLastCell).Row;
                int lastColumn = MySheet.Cells.SpecialCells(eX.XlCellType.xlCellTypeLastCell).Column;

                for (int i = 2; i <= lastRow; i++)
                {
                    YeuCauSanXuat nv = new YeuCauSanXuat();
                    string PROD_REQUEST_DATE, CODE_50, CODE_55, G_CODE, RIV_NO, PROD_REQUEST_QTY, CUST_CD, EMPL_NO, REMK, DELIVERY_DT, ext1, ext2, ext3, ext4, ext5, ext6, ext7, ext8, ext9, ext10, ext11, ext12, ext13, ext14, ext15, ext16, ext17, ext18, ext19, ext20;
;

                    PROD_REQUEST_DATE = MySheet.Range["A" + i].Value + "";
                    CODE_50 = MySheet.Range["B" + i].Value + "";
                    CODE_55 = MySheet.Range["C" + i].Value + "";
                    G_CODE = MySheet.Range["D" + i].Value + "";
                    RIV_NO = MySheet.Range["E" + i].Value + "";
                    PROD_REQUEST_QTY = MySheet.Range["F" + i].Value + "";
                    CUST_CD = MySheet.Range["G" + i].Value + "";
                    EMPL_NO = MySheet.Range["H" + i].Value + "";
                    REMK = MySheet.Range["I" + i].Value + "";
                    DELIVERY_DT = MySheet.Range["J" + i].Value + "";
                    ext1 = MySheet.Range["K" + i].Value + "";
                    ext2 = MySheet.Range["L" + i].Value + "";
                    ext3 = MySheet.Range["M" + i].Value + "";
                    ext4 = MySheet.Range["N" + i].Value + "";
                    ext5 = MySheet.Range["O" + i].Value + "";
                    ext6 = MySheet.Range["P" + i].Value + "";
                    ext7 = MySheet.Range["Q" + i].Value + "";
                    ext8 = MySheet.Range["R" + i].Value + "";
                    ext9 = MySheet.Range["S" + i].Value + "";
                    ext10 = MySheet.Range["T" + i].Value + "";
                    ext11 = MySheet.Range["U" + i].Value + "";
                    ext12 = MySheet.Range["V" + i].Value + "";
                    ext13 = MySheet.Range["W" + i].Value + "";
                    ext14 = MySheet.Range["X" + i].Value + "";
                    ext15 = MySheet.Range["Y" + i].Value + "";
                    ext16 = MySheet.Range["Z" + i].Value + "";
                    ext17 = MySheet.Range["AA" + i].Value + "";
                    ext18 = MySheet.Range["AB" + i].Value + "";
                    ext19 = MySheet.Range["AC" + i].Value + "";
                    ext20 = MySheet.Range["AD" + i].Value + "";




                    nv.PROD_REQUEST_DATE = PROD_REQUEST_DATE;
                    nv.CODE_50 = CODE_50;
                    nv.G_CODE = G_CODE;
                    nv.RIV_NO = RIV_NO;
                    nv.CODE_55 = CODE_55;
                    nv.PROD_REQUEST_QTY = PROD_REQUEST_QTY;
                    nv.CUST_CD = CUST_CD;
                    nv.EMPL_NO = EMPL_NO;
                    nv.REMK = REMK;
                    nv.DELIVERY_DT = DELIVERY_DT;
                    nv.ext1 = ext1;
                    nv.ext2 = ext2;
                    nv.ext3 = ext3;
                    nv.ext4 = ext4;
                    nv.ext5 = ext5;
                    nv.ext6 = ext6;
                    nv.ext7 = ext7;
                    nv.ext8 = ext8;
                    nv.ext9 = ext9;
                    nv.ext10 = ext10;
                    nv.ext11 = ext11;
                    nv.ext12 = ext12;
                    nv.ext13 = ext13;
                    nv.ext14 = ext14;
                    nv.ext15 = ext15;
                    nv.ext16 = ext16;
                    nv.ext17 = ext17;
                    nv.ext18 = ext18;
                    nv.ext19 = ext19;
                    nv.ext20 = ext20;
                    dsNV.Add(nv);
                }
                MyBook.Close(true);
                MyApp.Quit();
                releaseObject(MySheet);
                releaseObject(MyBook);
                releaseObject(MyApp);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return dsNV;
        }



        private static void copyAlltoClipboard(DataGridView dtgv)
        {
            dtgv.SelectAll();
            DataObject dataObj = dtgv.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }


        public static void exportReportToExcel(DataGridView dtgvobj, DataGridView dtgvobj2, DataGridView dtgvobj3, DataGridView dtgvobj4, DataGridView dtgvobj5, DataGridView dtgvobj6, DataGridView dtgvobj7, DataGridView dtgvobj8)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Documents (*.xlsx)|*.xlsx";
            sfd.FileName = "report.xlsx";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                // Copy DataGridView results to clipboard
                DataGridView[] dtgv = {dtgvobj,dtgvobj2,dtgvobj3,dtgvobj4,dtgvobj5, dtgvobj6 , dtgvobj7, dtgvobj8 };
                String[] title = {"1.WEEKLY PO BY CUSTOMER / 고객별 주차별 PO", "2.WEEKLY  PO BALANCE BY CUSTOMER/ 고객별 주차별 PO 잔량", "3. WEEKLY PO BY TYPE/ 제품군별 주차별 PO", "4. WEEKLY PO BALANCE BY TYPE/ 제품군별 주차별 PO 잔량", "5.WEEKLY FCST  BY CUSTOMER/ 고객별 주차별 FCST", "6. WEEKLY PO BY TYPE/ 제품군별 주차별 PO (SAMSUNG)", "7. WEEKLY PO BALANCE BY TYPE/ 제품군별 주차별 PO 잔량", "8. CUSTOMER PO BALANCE BY TYPE" };
                object misValue = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Excel.Application xlexcel = new Microsoft.Office.Interop.Excel.Application();

                xlexcel.DisplayAlerts = false;
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlexcel.Workbooks.Add(misValue);
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
               
                for (int i=0; i<8;i++)
                {
                    copyAlltoClipboard(dtgv[i]);
                    Microsoft.Office.Interop.Excel.Range last = xlWorkSheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                    int lastUsedRow = last.Row + 3;
                    Microsoft.Office.Interop.Excel.Range CR = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[lastUsedRow, 1];
                    CR.Select();
                    //xlWorkSheet.Range["A"+lastUsedRow].Value = title[i];
                    //last = xlWorkSheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                    xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

                    Microsoft.Office.Interop.Excel.Range CR2 = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[lastUsedRow, 1];
                    CR2.Select();
                    xlWorkSheet.Range["A" + lastUsedRow].Value = title[i];
                    xlWorkSheet.get_Range("A" + lastUsedRow).Select();
                }

               // Microsoft.Office.Interop.Excel.Range delRng = xlWorkSheet.get_Range("A:A").Cells;
               // delRng.Delete(Type.Missing);




                // Save the excel file under the captured location from the SaveFileDialog
                xlWorkBook.SaveAs(sfd.FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlexcel.DisplayAlerts = true;
                xlWorkBook.Close(true, misValue, misValue);
                xlexcel.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlexcel);

                // Clear Clipboard and DataGridView selection
                Clipboard.Clear();
                dtgvobj.ClearSelection();

                // Open the newly saved excel file
                if (File.Exists(sfd.FileName))
                    System.Diagnostics.Process.Start(sfd.FileName);
            }
        }


        public static void exportReportToExcel2(DataGridView dtgvobj, DataGridView dtgvobj2, DataGridView dtgvobj3, DataGridView dtgvobj4, DataGridView dtgvobj5, DataGridView dtgvobj6, DataGridView dtgvobj7, DataGridView dtgvobj8)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Documents (*.xlsx)|*.xlsx";
            sfd.FileName = "report.xlsx";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                // Copy DataGridView results to clipboard
                DataGridView[] dtgv = { dtgvobj, dtgvobj2, dtgvobj3, dtgvobj4, dtgvobj5, dtgvobj6, dtgvobj7, dtgvobj8 };
                String[] title = { "1.WEEKLY PO BY CUSTOMER / 고객별 주차별 PO", "2.WEEKLY  PO BALANCE BY CUSTOMER/ 고객별 주차별 PO 잔량", "3. WEEKLY PO BY TYPE/ 제품군별 주차별 PO", "4. WEEKLY PO BALANCE BY TYPE/ 제품군별 주차별 PO 잔량", "5.WEEKLY FCST  BY CUSTOMER/ 고객별 주차별 FCST", "6. WEEKLY PO BY TYPE/ 제품군별 주차별 PO (SAMSUNG)", "7. WEEKLY PO BALANCE BY TYPE/ 제품군별 주차별 PO 잔량", "8. CUSTOMER PO BALANCE BY TYPE" };
                object misValue = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Excel.Application xlexcel = new Microsoft.Office.Interop.Excel.Application();

                xlexcel.DisplayAlerts = false;
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlexcel.Workbooks.Add(misValue);
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);


                for (int i = 0; i < 8; i++)
                {
                    copyAlltoClipboard(dtgv[i]);
                    Microsoft.Office.Interop.Excel.Range last = xlWorkSheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                    int lastUsedRow = last.Row + 3;
                    Microsoft.Office.Interop.Excel.Range CR = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[lastUsedRow, 1];
                    CR.Select();
                    //xlWorkSheet.Range["A"+lastUsedRow].Value = title[i];
                    //last = xlWorkSheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                    xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

                    Microsoft.Office.Interop.Excel.Range CR2 = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[lastUsedRow, 1];
                    CR2.Select();
                    xlWorkSheet.Range["A" + lastUsedRow].Value = title[i];
                    xlWorkSheet.get_Range("A" + lastUsedRow).Select();
                }

                // Microsoft.Office.Interop.Excel.Range delRng = xlWorkSheet.get_Range("A:A").Cells;
                // delRng.Delete(Type.Missing);




                // Save the excel file under the captured location from the SaveFileDialog
                xlWorkBook.SaveAs(sfd.FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlexcel.DisplayAlerts = true;
                xlWorkBook.Close(true, misValue, misValue);
                xlexcel.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlexcel);

                // Clear Clipboard and DataGridView selection
                Clipboard.Clear();
                dtgvobj.ClearSelection();

                // Open the newly saved excel file
                if (File.Exists(sfd.FileName))
                    System.Diagnostics.Process.Start(sfd.FileName);
            }
        }



        public static void writeToExcelFile(DataGridView dtgvobj)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Documents (*.xlsx)|*.xlsx";
            sfd.FileName = "output.xlsx";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                // Copy DataGridView results to clipboard
                copyAlltoClipboard(dtgvobj);

                object misValue = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Excel.Application xlexcel = new Microsoft.Office.Interop.Excel.Application();

                xlexcel.DisplayAlerts = false; // Without this you will get two confirm overwrite prompts
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlexcel.Workbooks.Add(misValue);
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                // Format column D as text before pasting results, this was required for my data
                // Microsoft.Office.Interop.Excel.Range rng = xlWorkSheet.get_Range("D:D").Cells;
                //rng.NumberFormat = "@";

                // Paste clipboard results to worksheet range
                Microsoft.Office.Interop.Excel.Range CR = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[1, 1];
                CR.Select();
                xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

                // For some reason column A is always blank in the worksheet. ¯\_(ツ)_/¯
                // Delete blank column A and select cell A1

                Microsoft.Office.Interop.Excel.Range delRng = xlWorkSheet.get_Range("A:A").Cells;
                delRng.Delete(Type.Missing);
                xlWorkSheet.get_Range("A1").Select();


                // Save the excel file under the captured location from the SaveFileDialog
                xlWorkBook.SaveAs(sfd.FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlexcel.DisplayAlerts = true;
                xlWorkBook.Close(true, misValue, misValue);
                xlexcel.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlexcel);

                // Clear Clipboard and DataGridView selection
                Clipboard.Clear();
                dtgvobj.ClearSelection();

                // Open the newly saved excel file
                if (File.Exists(sfd.FileName))
                    System.Diagnostics.Process.Start(sfd.FileName);
            }
        }

        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                obj = null;
            }
            finally
            {
                GC.Collect();
            }

        }

    }


    public class DataConfig
    {
        private SqlConnection con;// khai báo biến connect
        //  public string strConnect = "Data Source=14.160.33.198; Initial Catalog=CMS_VINA;User ID=sa;Password=Cms6886;";

        //khởi tạo mặc định
        public DataConfig()
        {
            Connect();
        }

        //hàm kết nối csdl

        public string encode(string data)
        {
            string output = "";
            foreach (var c in data)
            {
                output+= (int)(c) + "_";
            }
            return output;
        }

        // 15_25_35_45_
        public string decode(string data)
        {
            string output = "";
            string[] words = data.Split('_');

            foreach (var word in words)
            {
                //MessageBox.Show(word);
                if(word !="")
                {
                    int ascii_code = int.Parse(word);
                    char character = (char)ascii_code;
                    string text = character.ToString();
                    output += text;
                }
                
            }
            return output;
        }
        private void Connect()
        {
            try
            {
                // string strConnect = "Data Source=14.160.33.198; Initial Catalog=CMS_VINA;User ID=sa;Password=Cms6886;" + " Pooling=false;";
                //string strConnect = "Data Source=14.160.33.94,3003; Initial Catalog=CMS_VINA;User ID=sa;Password=*11021201$;" + " Pooling=false;";
                //string strConnect = "Data Source=192.168.1.136,3003; Initial Catalog=CMS_VINA;User ID=sa;Password=*11021201$;" + " Pooling=false;";
               
                try
                {
                    string line;
                    // Read the file and display it line by line.  
                    string ipAddress="", port="", username="", password="";
                    System.IO.StreamReader file =
                        new System.IO.StreamReader("db.txt");
                    if ((line = file.ReadLine()) != null)
                    {
                        //MessageBox.Show("Line content: " + line);
                        ipAddress = line;
                    }
                    if ((line = file.ReadLine()) != null)
                    {
                        //MessageBox.Show("Line content: " + line);
                        port = line;
                    }
                    if ((line = file.ReadLine()) != null)
                    {
                        //MessageBox.Show("Line content: " + line);
                        username = decode(line);
                    }
                    if ((line = file.ReadLine()) != null)
                    {
                        //MessageBox.Show("Line content: " + line);
                        password = decode(line);
                    }
                    file.Close();

                    string strConnect = $"Data Source={ipAddress},{port}; Initial Catalog=CMS_VINA;User ID={username};Password={password};" + " Pooling=false;";




                    
                 

                                     //   MessageBox.Show(strConnect);
                    con = new SqlConnection(strConnect); //khởi tạo connect
                    if (con.State == ConnectionState.Open)//nếu kết nối đang mở thì ta đóng lại
                        con.Close(); // đóng kết nối
                    con.Open();// mở kết nối         

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Có lỗi !\n" + ex.ToString());
                }

                         
               
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi Kết nối : " + ex.ToString());
            }
        }

        //hàm getdata
        public DataTable GetData(string strSQL)
        {
            DataTable result = new DataTable();
            SqlDataAdapter da = new SqlDataAdapter(strSQL, con);
            da.Fill(result);
            return result;
            con.Close();
        }

        // thêm,sửa,xóa
        public int executeNoneQuery(string sql)
        {
            int result = 0;
            if (con.State == ConnectionState.Closed)
                con.Open();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = con;
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = sql;//câu lệnh truy vấn
            result = cmd.ExecuteNonQuery();
            return result;
        }

        //trả về 1 đối tượng nào đó
        public object executeScalarQuery(string sql)
        {
            object result = 0;
            if (con.State == ConnectionState.Closed)
                con.Open();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = con;
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = sql;//câu lệnh truy vấn
            result = cmd.ExecuteScalar();
            return result;
        }

    }



    public class ProductBLL
    {
      
        public int GetLastFcstWeekOfLastYear(int yearnum)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            int kq = 0;
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT MAX(FCSTWEEKNO) MAXFCSTWEEK FROM ZTBFCSTTB WHERE FCSTYEAR=" + yearnum;
            result = config.GetData(strQuery);
            if (result.Rows.Count > 0)
            {
                foreach (DataRow row in result.Rows)
                {
                    kq = int.Parse(row[0].ToString());
                }
            }
            else
            {
                kq = 0;
            }
            return kq;
        }

        public int checkRIV_NO(string G_CODE, string RIV_NO)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            int kq = 0;
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT * FROM M140 WHERE  G_CODE = '" + G_CODE +"'  AND RIV_NO = '" + RIV_NO + "'";
            result = config.GetData(strQuery);
            if (result.Rows.Count > 0)
            {
                kq = 1;
            }
            else
            {
                kq = 0;
            }
            return kq;
        }

        public string checkM100UseYN(string G_CODE)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string kq = "N";
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT USE_YN FROM M100 WHERE  G_CODE = '" + G_CODE + "'";
            result = config.GetData(strQuery);
            if (result.Rows.Count > 0)
            {
                kq = result.Rows[0]["USE_YN"].ToString();
            }
            else
            {
                kq = "N";
            }
            return kq;
        }


        public DataTable M100_insert(string values)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "INSERT INTO M100 (CTR_CD,G_CODE,G_NAME,CODE_12,SEQ_NO,REV_NO,CODE_33,CUST_CD,G_CODE_C,G_CODE_V,G_CODE_K,CODE_27,CODE_28,PRT_DRT,PRT_YN,PROD_PRINT_TIMES,PACK_DRT,ROLE_EA_QTY,G_WIDTH,G_LENGTH,G_R,G_C,G_LG,G_SG_L,G_SG_R,G_CG,REMK,USE_YN, INS_EMPL, INS_DATE, UPD_EMPL, UPD_DATE, PROD_PROJECT, PROD_MODEL, G_NAME_KD, DRAW_LINK, EQ1, EQ2, PROD_DIECUT_STEP, PD, KNIFE_TYPE, KNIFE_LIFECYCLE, KNIFE_PRICE, RPM, PIN_DISTANCE, PROCESS_TYPE, G_C_R, DESCR, PROD_MAIN_MATERIAL, PROD_TYPE,BANVE,NO_INSPECTION) VALUES " + values;
            result = config.GetData(strQuery);
            //MessageBox.Show(strQuery);
            return result;
        }
        public DataTable M100_update(string values)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "UPDATE M100 SET " + values;
            result = config.GetData(strQuery);
            return result;
        }

        public DataTable checktonkhofull_gcode(string G_CODE)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = $"SELECT M100.G_CODE, isnull(TONKIEM.INSPECT_BALANCE_QTY,0) AS CHO_KIEM, isnull(TONKIEM.WAIT_CS_QTY,0) AS CHO_CS_CHECK,isnull(TONKIEM.WAIT_SORTING_RMA,0) CHO_KIEM_RMA, isnull(TONKIEM.TOTAL_WAIT,0) AS TONG_TON_KIEM, isnull(BTP.BTP_QTY_EA,0) AS BTP, isnull(THANHPHAM.TONKHO,0) AS TON_TP, isnull(tbl_Block_table2.Block_Qty,0) AS BLOCK_QTY, (isnull(TONKIEM.TOTAL_WAIT,0) + isnull(BTP.BTP_QTY_EA,0)+ isnull(THANHPHAM.TONKHO,0) - isnull(tbl_Block_table2.Block_Qty,0)) AS GRAND_TOTAL_STOCK FROM M100 LEFT JOIN ( SELECT Product_MaVach, isnull([IN],0) AS NHAPKHO, isnull([OUT],0) AS XUATKHO, (isnull([IN],0)- isnull([OUT],0)) AS TONKHO FROM ( SELECT Product_Mavach, IO_Type, IO_Qty FROM tbl_InputOutput ) AS SourceTable PIVOT ( SUM(IO_Qty) FOR IO_Type IN ([IN], [OUT]) ) AS PivotTable ) AS THANHPHAM ON (THANHPHAM.Product_MaVach = M100.G_CODE) LEFT JOIN ( SELECT ZTB_WAIT_INSPECT.G_CODE, M100.G_NAME, M100.G_NAME_KD, SUM(INSPECT_BALANCE_QTY) AS INSPECT_BALANCE_QTY, SUM(WAIT_CS_QTY) AS WAIT_CS_QTY, SUM(WAIT_SORTING_RMA) AS WAIT_SORTING_RMA, SUM(INSPECT_BALANCE_QTY+ WAIT_CS_QTY+ WAIT_SORTING_RMA) AS TOTAL_WAIT FROM ZTB_WAIT_INSPECT JOIN M100 ON ( M100.G_CODE = ZTB_WAIT_INSPECT.G_CODE) WHERE UPDATE_DATE=CONVERT(date,GETDATE()) AND CALAMVIEC = 'DEM' GROUP BY ZTB_WAIT_INSPECT.G_CODE, M100.G_NAME, M100.G_NAME_KD) AS TONKIEM ON (THANHPHAM.Product_MaVach = TONKIEM.G_CODE) LEFT JOIN ( SELECT Product_MaVach, SUM(Block_Qty) AS Block_Qty from tbl_Block2 GROUP BY Product_MaVach ) AS tbl_Block_table2 ON (tbl_Block_table2.Product_MaVach= M100.G_CODE) LEFT JOIN ( SELECT ZTB_HALF_GOODS.G_CODE, M100.G_NAME, SUM(BTP_QTY_EA) AS BTP_QTY_EA FROM ZTB_HALF_GOODS JOIN M100 ON (M100.G_CODE = ZTB_HALF_GOODS.G_CODE) WHERE UPDATE_DATE = CONVERT(date,GETDATE()) GROUP BY ZTB_HALF_GOODS.G_CODE, M100.G_NAME) AS BTP ON (BTP.G_CODE = THANHPHAM.Product_MaVach) WHERE M100.G_CODE='{G_CODE}'";
            result = config.GetData(strQuery);
            return result;
        }

        public DataTable checkpobalance_gcode(string G_CODE)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = $"SELECT AA.G_CODE, (SUM(ZTBPOTable.PO_QTY)-SUM(AA.TotalDelivered)) As PO_BALANCE FROM (SELECT ZTBPOTable.EMPL_NO, ZTBPOTable.CUST_CD, ZTBPOTable.G_CODE, ZTBPOTable.PO_NO, isnull(SUM(ZTBDelivery.DELIVERY_QTY),0) AS TotalDelivered FROM ZTBPOTable LEFT JOIN ZTBDelivery ON (ZTBDelivery.CTR_CD = ZTBPOTable.CTR_CD AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) GROUP BY ZTBPOTable.CTR_CD,ZTBPOTable.EMPL_NO,ZTBPOTable.G_CODE,ZTBPOTable.CUST_CD,ZTBPOTable.PO_NO) AS AA JOIN ZTBPOTable ON ( AA.CUST_CD = ZTBPOTable.CUST_CD AND AA.G_CODE = ZTBPOTable.G_CODE AND AA.PO_NO = ZTBPOTable.PO_NO) WHERE AA.G_CODE='{G_CODE}' GROUP BY AA.G_CODE";
            result = config.GetData(strQuery);
            return result;
        }

        public DataTable checkfcst_gcode(string G_CODE)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = $"SELECT G_CODE, SUM(W1) AS W1, SUM(W2) AS W2, SUM(W3) AS W3, SUM(W4) AS W4, SUM(W5) AS W5, SUM(W6) AS W6, SUM(W7) AS W7, SUM(W8) AS W8 FROM ZTBFCSTTB WHERE FCSTYEAR = YEAR(GETDATE()) AND FCSTWEEKNO = (SELECT MAX(FCSTWEEKNO)FROM ZTBFCSTTB WHERE FCSTYEAR = YEAR(GETDATE()) ) AND G_CODE='{G_CODE}' GROUP BY G_CODE";
            result = config.GetData(strQuery);
            return result;
        }



        public DataTable pqc2_insert(string values)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "INSERT INTO ZTBPQC2TABLE (CTR_CD, PROCESS_LOT_NO, LINEQC_PIC, TIME1,TIME2,TIME3,TIME4,TIME5,TIME6,TIME7,TIME8,TIME9,TIME10,TIME11,TIME12,TIME13,TIME14,TIME15, CHECK1,CHECK2,CHECK3,CHECK4,CHECK5,CHECK6,CHECK7,CHECK8,CHECK9,CHECK10,CHECK11,CHECK12,CHECK13,CHECK14,CHECK15,REMARK) VALUES " + values;
            result = config.GetData(strQuery);
            return result;
        }

        public DataTable getLastProcessLotNo(string machine, string in_date)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = $"SELECT TOP 1 PROCESS_LOT_NO,SUBSTRING(PROCESS_LOT_NO,6,3) AS SEQ_NO, INS_DATE FROM P501 WHERE SUBSTRING(PROCESS_LOT_NO,1,2) = '{machine}' AND PROCESS_IN_DATE = '{in_date}' ORDER BY INS_DATE DESC";
            result = config.GetData(strQuery);
            return result;
        }



        public String getLastG_CODE_SEQ_NO(string dactinh, string phanloai)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string systemDate = "AAAA";
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT MAX(SEQ_NO) AS LAST_SEQ_NO FROM M100 WHERE CODE_12 = '" +dactinh+"' AND CODE_27='"+phanloai+"'";
            result = config.GetData(strQuery);
            
            if (result.Rows.Count > 0)
            {
                foreach (DataRow row in result.Rows)
                {
                    systemDate = row[0].ToString();
                }
            }
            else
            {
                systemDate = "ERROR";
            }

            return systemDate;
        }

        public String getlastver(string g_code)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            String systemDate = "AAAA";
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = $"SELECT TOP 1 REV_NO FROM M100 WHERE G_CODE LIKE '%{g_code}%' ORDER BY REV_NO DESC";
            result = config.GetData(strQuery);

            if (result.Rows.Count > 0)
            {
                foreach (DataRow row in result.Rows)
                {
                    systemDate = row[0].ToString();
                }
            }
            else
            {
                systemDate = "ERROR";
            }

            return systemDate;
        }

        public String getsystemDateTime()
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            String systemDate = "";
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT GETDATE() AS SYSTEM_DATE";
            result = config.GetData(strQuery);
            if (result.Rows.Count > 0)
            {
                foreach (DataRow row in result.Rows)
                {
                    systemDate = row[0].ToString();
                }
            }
            else
            {
                systemDate = "ERROR";
            }

            return systemDate;

        }


        public DataTable Material_History(string YCSX_NO, string chitiet)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "";
            if (chitiet == "chitiet")
            {
                strQuery = "DECLARE @YCSXNO AS VARCHAR(7) SET @YCSXNO = '" + YCSX_NO + "' SELECT ZZ.PROD_REQUEST_NO, ZZ.MIN_QTY , P400.PROD_REQUEST_QTY, M100.G_LENGTH, M100.G_LG, M100.G_C,  CONVERT(int,(ZZ.MIN_QTY/(M100.G_LENGTH + M100.G_LG))*M100.G_C*1000) AS QTY_DU_KIEN FROM ( SELECT @YCSXNO AS PROD_REQUEST_NO,isnull(MIN(YY.OUT_CFM_QTY),0) AS MIN_QTY  FROM (SELECT  XX.M_NAME, SUM(XX.OUT_CFM_QTY) AS OUT_CFM_QTY FROM ( SELECT O300.PROD_REQUEST_NO, O301.M_CODE, M090.M_NAME, M090.WIDTH_CD, O302.M_LOT_NO, O302.OUT_CFM_QTY FROM O300 JOIN O301 ON (O300.OUT_DATE = O301.OUT_DATE AND O301.OUT_NO = O300.OUT_NO) JOIN O302 ON (O302.OUT_DATE = O301.OUT_DATE AND O301.OUT_NO = O302.OUT_NO) JOIN M090 ON (M090.M_CODE = O301.M_CODE) WHERE O300.PROD_REQUEST_NO = @YCSXNO ) AS XX GROUP BY XX.M_NAME ) AS YY ) AS ZZ JOIN P400 ON (P400.PROD_REQUEST_NO = ZZ.PROD_REQUEST_NO) JOIN M100 ON (P400.G_CODE = M100.G_CODE)";
            }
            else if (chitiet == "chitiethon")
            {
                strQuery = "DECLARE @YCSXNO AS VARCHAR(7) SET @YCSXNO = '" + YCSX_NO + "' SELECT  XX.M_NAME, SUM(XX.OUT_CFM_QTY) AS OUT_CFM_QTY FROM ( SELECT O300.PROD_REQUEST_NO, O301.M_CODE, M090.M_NAME, M090.WIDTH_CD, O302.M_LOT_NO, O302.OUT_CFM_QTY FROM O300 JOIN O301 ON (O300.OUT_DATE = O301.OUT_DATE AND O301.OUT_NO = O300.OUT_NO) JOIN O302 ON (O302.OUT_DATE = O301.OUT_DATE AND O301.OUT_NO = O302.OUT_NO) JOIN M090 ON (M090.M_CODE = O301.M_CODE) WHERE O300.PROD_REQUEST_NO = @YCSXNO ) AS XX GROUP BY XX.M_NAME";
            }
            else if (chitiet == "chitiethonnua")
            {
                strQuery = "DECLARE @YCSXNO AS VARCHAR(7) SET @YCSXNO = '" + YCSX_NO + "' SELECT O300.PROD_REQUEST_NO, O301.M_CODE, M090.M_NAME, M090.WIDTH_CD, O302.M_LOT_NO, O302.OUT_CFM_QTY FROM O300 JOIN O301 ON (O300.OUT_DATE = O301.OUT_DATE AND O301.OUT_NO = O300.OUT_NO) JOIN O302 ON (O302.OUT_DATE = O301.OUT_DATE AND O301.OUT_NO = O302.OUT_NO) JOIN M090 ON (M090.M_CODE = O301.M_CODE) WHERE O300.PROD_REQUEST_NO = @YCSXNO";
            }
            else if (chitiet == "lieuinput")
            {
                strQuery = "SELECT O302.OUT_DATE,P501.M_LOT_NO, P500_A.M_CODE, M090.M_NAME, M090.WIDTH_CD, isnull(O302.OUT_CFM_QTY,0) AS OUT_CFM_QTY FROM P501 LEFT JOIN(SELECT DISTINCT PROD_REQUEST_NO, M_CODE, PROCESS_IN_DATE, PROCESS_IN_NO FROM P500) AS P500_A ON (P500_A.PROCESS_IN_DATE = P501.PROCESS_IN_DATE AND P500_A.PROCESS_IN_NO = P501.PROCESS_IN_NO) LEFT JOIN M090 ON P500_A.M_CODE = M090.M_CODE LEFT JOIN O302 ON (O302.M_LOT_NO = P501.M_LOT_NO) LEFT JOIN P400 ON (P400.PROD_REQUEST_NO = P500_A.PROD_REQUEST_NO) WHERE P500_A.PROD_REQUEST_NO='"+ YCSX_NO + "'";
            }

                result = config.GetData(strQuery);
            return result;

        }


        public DataTable chang_YCSXMANAGER_STATUS(string YCSX_NO, string status)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "UPDATE P400 SET YCSX_PENDING=" + status + "WHERE PROD_REQUEST_NO='" + YCSX_NO + "'";
            result = config.GetData(strQuery);
            return result;

        }



        public DataTable tra_YCSXMANAGER(string condition)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT P400.G_CODE, M100.G_NAME, M010.EMPL_NAME, M110.CUST_NAME_KD, P400.PROD_REQUEST_NO, P400.PROD_REQUEST_DATE, P400.PROD_REQUEST_QTY, isnull( INSPECT_BALANCE_TB.LOT_TOTAL_INPUT_QTY_EA, 0 ) AS LOT_TOTAL_INPUT_QTY_EA, isnull( INSPECT_BALANCE_TB.LOT_TOTAL_OUTPUT_QTY_EA, 0 ) AS LOT_TOTAL_OUTPUT_QTY_EA, isnull( INSPECT_BALANCE_TB.INSPECT_BALANCE, 0 ) AS INSPECT_BALANCE, ( CASE WHEN P400.YCSX_PENDING = 1 THEN (isnull(P400.PROD_REQUEST_QTY ,0)- isnull(INSPECT_BALANCE_TB.LOT_TOTAL_INPUT_QTY_EA,0)) WHEN P400.YCSX_PENDING = 0 THEN 0 END ) AS SHORTAGE_YCSX, ( CASE WHEN P400.YCSX_PENDING = 1 THEN 'PENDING' WHEN P400.YCSX_PENDING = 0 THEN 'CLOSED' END ) AS YCSX_PENDING, ( CASE WHEN P400.CODE_55 = '01' THEN 'Thong Thuong' WHEN P400.CODE_55 = '02' THEN 'SDI' WHEN P400.CODE_55 = '03' THEN 'GC' WHEN P400.CODE_55 = '04' THEN 'SAMPLE' END ) AS PHAN_LOAI, P400.REMK AS REMARK, P400.PO_TDYCSX, (P400.TKHO_TDYCSX+ P400.BTP_TDYCSX+ P400.CK_TDYCSX-  P400.BLOCK_TDYCSX) AS TOTAL_TKHO_TDYCSX, P400.TKHO_TDYCSX, P400.BTP_TDYCSX, P400.CK_TDYCSX, P400.BLOCK_TDYCSX, P400.FCST_TDYCSX, P400.W1,P400.W2,P400.W3,P400.W4,P400.W5,P400.W6,P400.W7,P400.W8, P400.PDUYET FROM P400 LEFT JOIN M100 ON (P400.G_CODE = M100.G_CODE) LEFT JOIN M010 ON (M010.EMPL_NO = P400.EMPL_NO) LEFT JOIN M110 ON (P400.CUST_CD = M110.CUST_CD) LEFT JOIN ( SELECT M010.EMPL_NAME, M110.CUST_NAME_KD, M100.G_CODE, M100.G_NAME, P400.PROD_REQUEST_NO, P400.PROD_REQUEST_QTY, INOUT.LOT_TOTAL_INPUT_QTY_EA, INOUT.LOT_TOTAL_OUTPUT_QTY_EA, INOUT.INSPECT_BALANCE FROM ( SELECT P400.PROD_REQUEST_NO, SUM(CC.LOT_TOTAL_INPUT_QTY_EA) AS LOT_TOTAL_INPUT_QTY_EA, SUM(CC.LOT_TOTAL_OUTPUT_QTY_EA) AS LOT_TOTAL_OUTPUT_QTY_EA, SUM(CC.INSPECT_BALANCE) AS INSPECT_BALANCE FROM ( SELECT AA.PROCESS_LOT_NO, AA.LOT_TOTAL_QTY_KG, AA.LOT_TOTAL_INPUT_QTY_EA, isnull(BB.LOT_TOTAL_OUTPUT_QTY_EA, 0) AS LOT_TOTAL_OUTPUT_QTY_EA, ( AA.LOT_TOTAL_INPUT_QTY_EA - isnull(BB.LOT_TOTAL_OUTPUT_QTY_EA, 0) ) AS INSPECT_BALANCE FROM ( SELECT PROCESS_LOT_NO, SUM(INPUT_QTY_EA) As LOT_TOTAL_INPUT_QTY_EA, SUM(INPUT_QTY_KG) AS LOT_TOTAL_QTY_KG FROM ZTBINSPECTINPUTTB GROUP BY PROCESS_LOT_NO ) AS AA LEFT JOIN ( SELECT PROCESS_LOT_NO, SUM(OUTPUT_QTY_EA) As LOT_TOTAL_OUTPUT_QTY_EA FROM ZTBINSPECTOUTPUTTB GROUP BY PROCESS_LOT_NO ) AS BB ON ( AA.PROCESS_LOT_NO = BB.PROCESS_LOT_NO ) ) AS CC LEFT JOIN P501 ON ( CC.PROCESS_LOT_NO = P501.PROCESS_LOT_NO ) LEFT JOIN ( SELECT DISTINCT PROD_REQUEST_NO, PROCESS_IN_DATE, PROCESS_IN_NO FROM P500 ) AS P500_A ON ( P500_A.PROCESS_IN_DATE = P501.PROCESS_IN_DATE AND P500_A.PROCESS_IN_NO = P501.PROCESS_IN_NO ) LEFT JOIN P400 ON ( P500_A.PROD_REQUEST_NO = P400.PROD_REQUEST_NO ) GROUP BY P400.PROD_REQUEST_NO ) AS INOUT LEFT JOIN P400 ON ( INOUT.PROD_REQUEST_NO = P400.PROD_REQUEST_NO ) LEFT JOIN M110 ON (M110.CUST_CD = P400.CUST_CD) LEFT JOIN M100 ON (M100.G_CODE = P400.G_CODE) LEFT JOIN M010 ON (M010.EMPL_NO = P400.EMPL_NO) ) AS INSPECT_BALANCE_TB ON ( INSPECT_BALANCE_TB.PROD_REQUEST_NO = P400.PROD_REQUEST_NO )" + condition + " ORDER BY P400.INS_DATE DESC";
            result = config.GetData(strQuery);
            return result;

        }



        public DataTable report_inspection_all_balance_data_YCSX(string condition)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            //string strQuery = "SELECT M010.EMPL_NAME, M110.CUST_NAME_KD, M100.G_CODE, M100.G_NAME, P400.PROD_REQUEST_NO, P400.PROD_REQUEST_DATE,P400.PROD_REQUEST_QTY,  INOUT.LOT_TOTAL_INPUT_QTY_EA, INOUT.LOT_TOTAL_OUTPUT_QTY_EA, INOUT.INSPECT_BALANCE, isnull(ZTBINSPECTNGTB_A.INSPECT_TOTAL_QTY,0) AS  INSPECT_TOTAL_QTY, isnull(ZTBINSPECTNGTB_A.INSPECT_OK_QTY,0) AS INSPECT_OK_QTY , isnull(ZTBINSPECTNGTB_A.LOSS_AND_NG_AND_MARKING_QTY,0) AS LOSS_AND_NG_AND_MARKING_QTY, isnull(ZTBINSPECTNGTB_A.LOSS_THEM_TUI_QTY,0) AS LOSS_THEM_TUI_QTY FROM ( SELECT  P400.PROD_REQUEST_NO ,  SUM(CC.LOT_TOTAL_INPUT_QTY_EA) AS LOT_TOTAL_INPUT_QTY_EA , SUM(CC.LOT_TOTAL_OUTPUT_QTY_EA) AS LOT_TOTAL_OUTPUT_QTY_EA , SUM(CC.INSPECT_BALANCE) AS  INSPECT_BALANCE FROM( 	SELECT  AA.PROCESS_LOT_NO,  AA.LOT_TOTAL_QTY_KG, AA.LOT_TOTAL_INPUT_QTY_EA, isnull(BB.LOT_TOTAL_OUTPUT_QTY_EA,0) AS LOT_TOTAL_OUTPUT_QTY_EA, ( AA.LOT_TOTAL_INPUT_QTY_EA- isnull(BB.LOT_TOTAL_OUTPUT_QTY_EA,0)) AS INSPECT_BALANCE FROM 	(SELECT PROCESS_LOT_NO, SUM(INPUT_QTY_EA) As LOT_TOTAL_INPUT_QTY_EA, SUM(INPUT_QTY_KG) AS LOT_TOTAL_QTY_KG  FROM ZTBINSPECTINPUTTB GROUP BY PROCESS_LOT_NO) AS AA 	LEFT JOIN 	(SELECT PROCESS_LOT_NO, SUM(OUTPUT_QTY_EA) As LOT_TOTAL_OUTPUT_QTY_EA  FROM ZTBINSPECTOUTPUTTB GROUP BY PROCESS_LOT_NO) AS BB 	ON (AA.PROCESS_LOT_NO = BB.PROCESS_LOT_NO) ) AS CC JOIN P501 ON (CC.PROCESS_LOT_NO = P501.PROCESS_LOT_NO) JOIN (SELECT DISTINCT PROD_REQUEST_NO, PROCESS_IN_DATE, PROCESS_IN_NO FROM P500) AS P500_A ON (P500_A.PROCESS_IN_DATE = P501.PROCESS_IN_DATE AND P500_A.PROCESS_IN_NO = P501.PROCESS_IN_NO) JOIN P400 ON (P500_A.PROD_REQUEST_NO = P400.PROD_REQUEST_NO) GROUP BY P400.PROD_REQUEST_NO ) AS INOUT JOIN P400 ON (INOUT.PROD_REQUEST_NO = P400.PROD_REQUEST_NO) JOIN M110 ON (M110.CUST_CD = P400.CUST_CD)  JOIN M100 ON (M100.G_CODE = P400.G_CODE) JOIN M010 ON (M010.EMPL_NO = P400.EMPL_NO) LEFT JOIN ( 	SELECT  P500_AA.PROD_REQUEST_NO, SUM(ERR1+ERR2+ERR2) AS TOTAL_LOSS, 	SUM(ERR4+ERR5+ERR6+ERR7+ERR8+ERR9+ERR10+ERR11+ERR12+ERR13+ERR14+ERR15+ERR16+ERR17+ERR18+ERR19+ERR20+ERR21+ERR22+ERR23+ERR24+ERR25+ERR26+ERR27+ERR28+ERR29+ERR30+ERR31) AS INSPECT_TOTAL_NG_QTY, 	SUM(INSPECT_TOTAL_QTY) AS INSPECT_TOTAL_QTY , 	SUM(INSPECT_OK_QTY) AS INSPECT_OK_QTY, 	SUM(ERR2+ERR3+ERR4+ERR5+ERR6+ERR7+ERR8+ERR9+ERR10+ERR11+ERR12+ERR13+ERR14+ERR15+ERR16+ERR17+ERR18+ERR19+ERR20+ERR21+ERR22+ERR23+ERR24+ERR25+ERR26+ERR27+ERR28+ERR29+ERR30+ERR31+ERR32) AS LOSS_AND_NG_AND_MARKING_QTY, 	SUM(ERR1) AS LOSS_THEM_TUI_QTY 	FROM ZTBINSPECTNGTB 	LEFT JOIN P501 ON (P501.PROCESS_LOT_NO = ZTBINSPECTNGTB.PROCESS_LOT_NO) 	LEFT JOIN (SELECT DISTINCT PROD_REQUEST_NO, PROCESS_IN_DATE, PROCESS_IN_NO FROM P500) AS P500_AA ON  (P500_AA.PROCESS_IN_DATE = P501.PROCESS_IN_DATE AND P500_AA.PROCESS_IN_NO = P501.PROCESS_IN_NO) 	GROUP BY P500_AA.PROD_REQUEST_NO ) AS ZTBINSPECTNGTB_A ON (ZTBINSPECTNGTB_A.PROD_REQUEST_NO = INOUT.PROD_REQUEST_NO)" + condition;
            string strQuery = "SELECT M010.EMPL_NAME AS PIC_KD,M110.CUST_NAME_KD, M100.G_CODE, M100.G_NAME, M100.G_NAME_KD, INPUTTB.PROD_REQUEST_NO, P400.PROD_REQUEST_DATE, P400.PROD_REQUEST_QTY, INPUTTB.INPUT_QTY AS LOT_TOTAL_INPUT_QTY_EA, isnull(OUTPUTTB.OUTPUT_QTY,0) AS LOT_TOTAL_OUTPUT_QTY_EA, isnull(INSPECTTABLE.DA_KIEM_TRA,0) AS DA_KIEM_TRA, isnull(INSPECTTABLE.OK_QTY,0) AS OK_QTY, isnull(INSPECTTABLE.LOSS_NG_QTY,0) AS LOSS_NG_QTY, (isnull(INPUTTB.INPUT_QTY,0) -  isnull(INSPECTTABLE.DA_KIEM_TRA,0)) AS INSPECT_BALANCE FROM  (SELECT PROD_REQUEST_NO, SUM(INPUT_QTY_EA) AS INPUT_QTY FROM ZTBINSPECTINPUTTB  GROUP BY PROD_REQUEST_NO) AS INPUTTB LEFT JOIN (SELECT PROD_REQUEST_NO, SUM(OUTPUT_QTY_EA) AS OUTPUT_QTY FROM ZTBINSPECTOUTPUTTB  GROUP BY PROD_REQUEST_NO) AS OUTPUTTB  ON (INPUTTB.PROD_REQUEST_NO = OUTPUTTB.PROD_REQUEST_NO)  LEFT JOIN (SELECT PROD_REQUEST_NO, SUM(INSPECT_TOTAL_QTY) AS DA_KIEM_TRA,SUM(INSPECT_OK_QTY) AS OK_QTY,SUM(INSPECT_TOTAL_QTY- INSPECT_OK_QTY) AS LOSS_NG_QTY FROM ZTBINSPECTNGTB GROUP BY PROD_REQUEST_NO) AS INSPECTTABLE ON (INPUTTB.PROD_REQUEST_NO = INSPECTTABLE.PROD_REQUEST_NO)  LEFT JOIN P400 ON (P400.PROD_REQUEST_NO = INPUTTB.PROD_REQUEST_NO)  LEFT JOIN M100 ON (P400.G_CODE = M100.G_CODE)  LEFT JOIN M010 ON (M010.EMPL_NO = P400.EMPL_NO)  LEFT JOIN M110 ON (M110.CUST_CD = P400.CUST_CD)" + condition;
            result = config.GetData(strQuery);
            return result;

        }



        public DataTable report_inspection_all_output_data_YCSX(string condition)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();            
            string strQuery = "SELECT M110.CUST_NAME_KD, M010.EMPL_NAME, P400.PROD_REQUEST_NO,M100.G_CODE, M100.G_NAME, P400.PROD_REQUEST_QTY, ZTBINSPECTOUTPUTTB_YCSX.OUTPUT_QTY_EA FROM ( SELECT P400.PROD_REQUEST_NO,  SUM(ZTBINSPECTOUTPUTTB.OUTPUT_QTY_EA) As OUTPUT_QTY_EA FROM ZTBINSPECTOUTPUTTB JOIN P501 ON (ZTBINSPECTOUTPUTTB.PROCESS_LOT_NO = P501.PROCESS_LOT_NO) JOIN (SELECT DISTINCT PROD_REQUEST_NO, PROCESS_IN_DATE, PROCESS_IN_NO FROM P500) AS P500_A ON (P500_A.PROCESS_IN_DATE = P501.PROCESS_IN_DATE AND P500_A.PROCESS_IN_NO = P501.PROCESS_IN_NO) JOIN P400 ON (P500_A.PROD_REQUEST_NO = P400.PROD_REQUEST_NO) GROUP BY P400.PROD_REQUEST_NO ) AS ZTBINSPECTOUTPUTTB_YCSX JOIN P400 ON (P400.PROD_REQUEST_NO = ZTBINSPECTOUTPUTTB_YCSX.PROD_REQUEST_NO) JOIN M110 ON (M110.CUST_CD = P400.CUST_CD) JOIN M100 ON (M100.G_CODE = P400.G_CODE) JOIN M010 ON (M010.EMPL_NO = P400.EMPL_NO) "+condition;
            result = config.GetData(strQuery);
            return result;

        }



        public DataTable report_inspection_all_input_data_YCSX(string condition)
        {

            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT M010.EMPL_NAME, M110.CUST_NAME_KD, ZTBINSPECTINPUTTB_YCSX.PROD_REQUEST_NO, P400.PROD_REQUEST_QTY, M100.G_CODE, M100.G_NAME, ZTBINSPECTINPUTTB_YCSX.INPUT_QTY_EA, ZTBINSPECTINPUTTB_YCSX.INPUT_QTY_KG FROM (SELECT P400.PROD_REQUEST_NO, SUM(ZTBINSPECTINPUTTB.INPUT_QTY_EA) AS INPUT_QTY_EA, SUM(ZTBINSPECTINPUTTB.INPUT_QTY_KG) AS INPUT_QTY_KG FROM ZTBINSPECTINPUTTB  JOIN P501 ON (ZTBINSPECTINPUTTB.PROCESS_LOT_NO = P501.PROCESS_LOT_NO) JOIN (SELECT DISTINCT PROD_REQUEST_NO, PROCESS_IN_DATE, PROCESS_IN_NO FROM P500) AS P500_A ON (P500_A.PROCESS_IN_DATE = P501.PROCESS_IN_DATE AND P500_A.PROCESS_IN_NO = P501.PROCESS_IN_NO) JOIN P400 ON (P500_A.PROD_REQUEST_NO = P400.PROD_REQUEST_NO) JOIN M110 ON (M110.CUST_CD = P400.CUST_CD) JOIN M100 ON (M100.G_CODE = P400.G_CODE) JOIN M010 ON (M010.EMPL_NO = ZTBINSPECTINPUTTB.EMPL_NO) GROUP BY P400.PROD_REQUEST_NO) AS ZTBINSPECTINPUTTB_YCSX JOIN P400 ON (ZTBINSPECTINPUTTB_YCSX.PROD_REQUEST_NO = P400.PROD_REQUEST_NO) JOIN M110 ON (M110.CUST_CD = P400.CUST_CD) JOIN M100 ON (M100.G_CODE = P400.G_CODE) JOIN M010 ON (M010.EMPL_NO = P400.EMPL_NO)" + condition;
            result = config.GetData(strQuery);
            return result;

        }


        public DataTable report_inspection_tra_NG()
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT  ZTBINSPECTNGTB.INSPECT_ID, P400.PROD_REQUEST_NO, M100.G_NAME, M100.G_CODE, M100.PROD_TYPE, P501.M_LOT_NO, M090.M_NAME, M090.WIDTH_CD, O302.OUT_CFM_QTY, M010.EMPL_NAME AS INSPECTOR, M010_A.EMPL_NAME AS LINEQC, ZTBINSPECTNGTB.EMPL_NO,ZTBINSPECTNGTB.PROCESS_LOT_NO,INSPECT_DATETIME, INSPECT_START_TIME, INSPECT_FINISH_TIME, FACTORY,LINEQC_PIC,MACHINE_NO,INSPECT_TOTAL_QTY,INSPECT_OK_QTY, (ERR1+ERR2+ERR3) AS INSPECT_TOTAL_LOSS_QTY, (ERR4+ERR5+ERR6+ERR7+ERR8+ERR9+ERR10+ERR11+ERR12+ERR13+ERR14+ERR15+ERR16+ERR17+ERR18+ERR19+ERR20+ERR21+ERR22+ERR23+ERR24+ERR25+ERR26+ERR27+ERR28+ERR29+ERR30+ERR31) AS INSPECT_TOTAL_NG_QTY, (ERR4+ERR5+ERR6+ERR7+ERR8+ERR9+ERR10) AS MATERIAL_NG_QTY, (ERR11+ERR12+ERR13+ERR14+ERR15+ERR16+ERR17+ERR18+ERR19+ERR20+ERR21+ERR22+ERR23+ERR24+ERR25+ERR26+ERR27+ERR28+ERR29+ERR30+ERR31) AS PROCESS_NG_QTY, ZTBPOTable_A.PROD_PRICE,ERR1, ERR2, ERR3,ERR4,ERR5,ERR6,ERR7,ERR8,ERR9,ERR10,ERR11,ERR12,ERR13,ERR14,ERR15,ERR16,ERR17,ERR18,ERR19,ERR20,ERR21,ERR22,ERR23,ERR24,ERR25,ERR26,ERR27,ERR28,ERR29,ERR30,ERR31,ERR32 FROM ZTBINSPECTNGTB JOIN P501 ON (P501.PROCESS_LOT_NO = ZTBINSPECTNGTB.PROCESS_LOT_NO) JOIN (SELECT DISTINCT PROD_REQUEST_NO, PROCESS_IN_DATE, PROCESS_IN_NO FROM P500) AS P500_A ON (P500_A.PROCESS_IN_DATE = P501.PROCESS_IN_DATE AND P500_A.PROCESS_IN_NO = P501.PROCESS_IN_NO) JOIN P400 ON (P500_A.PROD_REQUEST_NO = P400.PROD_REQUEST_NO) JOIN M110 ON (M110.CUST_CD = P400.CUST_CD) JOIN M100 ON (M100.G_CODE = P400.G_CODE) JOIN M010 ON (M010.EMPL_NO = ZTBINSPECTNGTB.EMPL_NO) JOIN (SELECT EMPL_NAME, EMPL_NO FROM M010) AS M010_A ON M010_A.EMPL_NO = ZTBINSPECTNGTB.LINEQC_PIC JOIN (SELECT DISTINCT G_CODE, MIN(PROD_PRICE) AS PROD_PRICE FROM ZTBPOTable GROUP BY G_CODE) AS ZTBPOTable_A ON (ZTBPOTable_A.G_CODE = M100.G_CODE) JOIN O302 ON (O302.M_LOT_NO = P501.M_LOT_NO) JOIN M090 ON (M090.M_CODE = O302.M_CODE) ORDER BY INSPECT_ID DESC";
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable report_inspection_insert_NG(string values)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "INSERT INTO ZTBINSPECTNGTB (CTR_CD,INSPECT_START_TIME, INSPECT_FINISH_TIME,EMPL_NO,PROCESS_LOT_NO, INSPECT_DATETIME, FACTORY, LINEQC_PIC, MACHINE_NO, INSPECT_TOTAL_QTY, INSPECT_OK_QTY,ERR1,ERR2,ERR3,ERR4,ERR5,ERR6,ERR7,ERR8,ERR9,ERR10,ERR11,ERR12,ERR13,ERR14,ERR15,ERR16,ERR17,ERR18,ERR19,ERR20,ERR21,ERR22,ERR23,ERR24,ERR25,ERR26,ERR27,ERR28,ERR29,ERR30,ERR31,ERR32) VALUES " + values;
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable report_inspection_insert_output(string values)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "INSERT INTO ZTBINSPECTOUTPUTTB (CTR_CD, EMPL_NO,PROCESS_LOT_NO,OUTPUT_DATETIME,OUTPUT_QTY_EA, REMARK) VALUES" + values;
            result = config.GetData(strQuery);
            return result;

        }


        public DataTable report_inspection_insert_input(string values)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "INSERT INTO ZTBINSPECTINPUTTB (CTR_CD, EMPL_NO,PROCESS_LOT_NO,INPUT_DATETIME,INPUT_QTY_EA,INPUT_QTY_KG, REMARK) VALUES" + values;
            result = config.GetData(strQuery);
            return result;

        }


        public DataTable report_inspection_check_lot_exist(string condition)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT PROCESS_LOT_NO FROM ZTBINSPECTINPUTTB WHERE PROCESS_LOT_NO='" + condition + "'";
            result = config.GetData(strQuery);
            return result;

        }


        public String report_inspection_check_EMPLNAME(String EMPL_NO)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            String empl_name = "";
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT EMPL_NAME FROM M010 WHERE EMPL_NO = '" + EMPL_NO + "'";
            result = config.GetData(strQuery);
            if (result.Rows.Count > 0)
            {
                foreach (DataRow row in result.Rows)
                {
                    empl_name = row[0].ToString();
                }
            }
            else
            {
                empl_name = "KHÔNG TỒN TẠI MÃ NHÂN VIÊN NÀY";
            }
            return empl_name;
        }

        public DataTable check_M_NAME(String M_LOT_NO)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            String empl_name = "";
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = $"SELECT M090.M_NAME FROM I222 JOIN M090 ON (M090.M_CODE=  I222.M_CODE) WHERE I222.M_LOT_NO ='{M_LOT_NO}'";
            result = config.GetData(strQuery);
            return result;
           
        }

        public String report_inspection_check_prod_date(String lotsx)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            String prod_date = "";
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT INS_DATE FROM P501 WHERE PROCESS_LOT_NO = '" + lotsx + "'";
            result = config.GetData(strQuery);
            if (result.Rows.Count > 0)
            {
                foreach (DataRow row in result.Rows)
                {
                    prod_date = row[0].ToString();
                }
            }
            else
            {
                prod_date = "KHÔNG TỒN TẠI LOT";
            }
            return prod_date;
        }

        public DataTable report_inspection_check_lot_no(string condition)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT G_NAME FROM P501 JOIN (SELECT DISTINCT PROD_REQUEST_NO, PROCESS_IN_DATE, PROCESS_IN_NO FROM P500) AS P500_A ON (P500_A.PROCESS_IN_DATE = P501.PROCESS_IN_DATE AND P500_A.PROCESS_IN_NO = P501.PROCESS_IN_NO) JOIN P400 ON (P400.PROD_REQUEST_NO = P500_A.PROD_REQUEST_NO) JOIN M100 ON (P400.G_CODE =M100.G_CODE) WHERE P501.PROCESS_LOT_NO='" + condition+"'";
            result = config.GetData(strQuery);
            return result;

        }


        public DataTable report_inspection_all_NG_data(string condition)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT ZTBINSPECTNGTB.INSPECT_ID, CONCAT(datepart(YEAR,INSPECT_START_TIME),'_',datepart(ISO_WEEK,DATEADD(day,2,INSPECT_START_TIME))) AS YEAR_WEEK, M110.CUST_NAME_KD,ZTBINSPECTNGTB.PROD_REQUEST_NO,M100.G_NAME_KD,M100.G_NAME,ZTBINSPECTNGTB.G_CODE,M100.PROD_TYPE,ZTBINSPECTNGTB.M_LOT_NO,M090.M_NAME,M090.WIDTH_CD,ZTBINSPECTNGTB.EMPL_NO AS INSPECTOR,ZTBINSPECTNGTB.LINEQC_PIC AS LINEQC,ZTBINSPECTNGTB.PROD_PIC,M100.CODE_33 AS UNIT ,ZTBINSPECTNGTB.PROCESS_LOT_NO,ZTBINSPECTNGTB.PROCESS_IN_DATE,ZTBINSPECTNGTB.INSPECT_DATETIME, ZTBINSPECTNGTB.INSPECT_START_TIME,ZTBINSPECTNGTB.INSPECT_FINISH_TIME,ZTBINSPECTNGTB.FACTORY,ZTBINSPECTNGTB.LINEQC_PIC,ZTBINSPECTNGTB.MACHINE_NO,ZTBINSPECTNGTB.INSPECT_TOTAL_QTY,ZTBINSPECTNGTB.INSPECT_OK_QTY,CAST(INSPECT_TOTAL_QTY AS float)/(CAST(DATEDIFF(MINUTE, ZTBINSPECTNGTB.INSPECT_START_TIME,ZTBINSPECTNGTB.INSPECT_FINISH_TIME) AS float) / CAST(60 as float) )  AS INSPECT_SPEED,(ERR1+ERR2+ERR3) AS INSPECT_TOTAL_LOSS_QTY, (ERR4+ERR5+ERR6+ERR7+ERR8+ERR9+ERR10+ERR11+ERR12+ERR13+ERR14+ERR15+ERR16+ERR17+ERR18+ERR19+ERR20+ERR21+ERR22+ERR23+ERR24+ERR25+ERR26+ERR27+ERR28+ERR29+ERR30+ERR31) AS INSPECT_TOTAL_NG_QTY, (ERR4+ERR5+ERR6+ERR7+ERR8+ERR9+ERR10+ERR11) AS MATERIAL_NG_QTY, (ERR12+ERR13+ERR14+ERR15+ERR16+ERR17+ERR18+ERR19+ERR20+ERR21+ERR22+ERR23+ERR24+ERR25+ERR26+ERR27+ERR28+ERR29+ERR30+ERR31) AS PROCESS_NG_QTY,M100.PROD_LAST_PRICE AS PROD_PRICE,ZTBINSPECTNGTB.ERR1,ZTBINSPECTNGTB.ERR2,ZTBINSPECTNGTB.ERR3,ZTBINSPECTNGTB.ERR4,ZTBINSPECTNGTB.ERR5,ZTBINSPECTNGTB.ERR6,ZTBINSPECTNGTB.ERR7,ZTBINSPECTNGTB.ERR8,ZTBINSPECTNGTB.ERR9,ZTBINSPECTNGTB.ERR10,ZTBINSPECTNGTB.ERR11,ZTBINSPECTNGTB.ERR12,ZTBINSPECTNGTB.ERR13,ZTBINSPECTNGTB.ERR14,ZTBINSPECTNGTB.ERR15,ZTBINSPECTNGTB.ERR16,ZTBINSPECTNGTB.ERR17,ZTBINSPECTNGTB.ERR18,ZTBINSPECTNGTB.ERR19,ZTBINSPECTNGTB.ERR20,ZTBINSPECTNGTB.ERR21,ZTBINSPECTNGTB.ERR22,ZTBINSPECTNGTB.ERR23,ZTBINSPECTNGTB.ERR24,ZTBINSPECTNGTB.ERR25,ZTBINSPECTNGTB.ERR26,ZTBINSPECTNGTB.ERR27,ZTBINSPECTNGTB.ERR28,ZTBINSPECTNGTB.ERR29,ZTBINSPECTNGTB.ERR30,ZTBINSPECTNGTB.ERR31,ZTBINSPECTNGTB.ERR32, ZTBINSPECTNGTB.CNDB_ENCODES  FROM ZTBINSPECTNGTB  LEFT JOIN M110 ON (ZTBINSPECTNGTB.CUST_CD = M110.CUST_CD)  LEFT JOIN M100 ON (ZTBINSPECTNGTB.G_CODE = M100.G_CODE) LEFT JOIN M090 ON(ZTBINSPECTNGTB.M_CODE = M090.M_CODE) " + condition + " ORDER BY INSPECT_ID DESC";
            result = config.GetData(strQuery);
            return result;

        }


        public DataTable report_inspection_all_balance_data(string condition)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT M110.CUST_NAME_KD, M100.G_NAME, M100.PROD_TYPE, M100.G_CODE, P400.PROD_REQUEST_DATE, P400.PROD_REQUEST_NO, P400.PROD_REQUEST_QTY, P501.M_LOT_NO, isnull(M090.M_NAME,'NO OUTPUT') AS M_NAME, isnull(M090.WIDTH_CD,0) As WIDTH_CD, isnull(O302.OUT_CFM_QTY,0) AS OUT_CFM_QTY, CC.PROCESS_LOT_NO, CC.LOT_TOTAL_INPUT_QTY_EA, CC.LOT_TOTAL_OUTPUT_QTY_EA, CC.INSPECT_BALANCE , isnull(ZTBINSPECTNGTB_A.TOTAL_LOSS,0) AS TOTAL_LOSS, isnull(ZTBINSPECTNGTB_A.INSPECT_TOTAL_NG_QTY,0) AS TOTAL_NG_QTY FROM ( 	SELECT  AA.PROCESS_LOT_NO,  AA.LOT_TOTAL_QTY_KG, AA.LOT_TOTAL_INPUT_QTY_EA, isnull(BB.LOT_TOTAL_OUTPUT_QTY_EA,0) AS LOT_TOTAL_OUTPUT_QTY_EA, ( AA.LOT_TOTAL_INPUT_QTY_EA- isnull(BB.LOT_TOTAL_OUTPUT_QTY_EA,0)) AS INSPECT_BALANCE FROM 	(SELECT PROCESS_LOT_NO, SUM(INPUT_QTY_EA) As LOT_TOTAL_INPUT_QTY_EA, SUM(INPUT_QTY_KG) AS LOT_TOTAL_QTY_KG  FROM ZTBINSPECTINPUTTB GROUP BY PROCESS_LOT_NO) AS AA 	LEFT JOIN 	(SELECT PROCESS_LOT_NO, SUM(OUTPUT_QTY_EA) As LOT_TOTAL_OUTPUT_QTY_EA  FROM ZTBINSPECTOUTPUTTB GROUP BY PROCESS_LOT_NO) AS BB 	ON (AA.PROCESS_LOT_NO = BB.PROCESS_LOT_NO) ) AS CC JOIN P501 ON (CC.PROCESS_LOT_NO = P501.PROCESS_LOT_NO) JOIN (SELECT DISTINCT PROD_REQUEST_NO, PROCESS_IN_DATE, PROCESS_IN_NO FROM P500) AS P500_A ON (P500_A.PROCESS_IN_DATE = P501.PROCESS_IN_DATE AND P500_A.PROCESS_IN_NO = P501.PROCESS_IN_NO) JOIN P400 ON (P500_A.PROD_REQUEST_NO = P400.PROD_REQUEST_NO) JOIN M110 ON (M110.CUST_CD = P400.CUST_CD)  JOIN M100 ON (M100.G_CODE = P400.G_CODE) LEFT JOIN (SELECT  PROCESS_LOT_NO, SUM(ERR1+ERR2+ERR2) AS TOTAL_LOSS, SUM(ERR4+ERR5+ERR6+ERR7+ERR8+ERR9+ERR10+ERR11+ERR12+ERR13+ERR14+ERR15+ERR16+ERR17+ERR18+ERR19+ERR20+ERR21+ERR22+ERR23+ERR24+ERR25+ERR26+ERR27+ERR28+ERR29+ERR30+ERR31) AS INSPECT_TOTAL_NG_QTY FROM ZTBINSPECTNGTB GROUP BY PROCESS_LOT_NO) AS ZTBINSPECTNGTB_A ON (ZTBINSPECTNGTB_A.PROCESS_LOT_NO = CC.PROCESS_LOT_NO) LEFT JOIN O302 ON (O302.M_LOT_NO = P501.M_LOT_NO) LEFT JOIN M090 ON (M090.M_CODE = O302.M_CODE)" + condition;
            result = config.GetData(strQuery);
            return result;

        }



        public DataTable report_inspection_all_output_data(string condition)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";

            //SELECT ZTBINSPECTOUTPUTTB.INSPECT_OUTPUT_ID, M110.CUST_NAME_KD,M010.EMPL_NAME, M100.G_CODE, M100.G_NAME, ZTBINSPECTOUTPUTTB.PROCESS_LOT_NO, P400.PROD_REQUEST_NO, P400.PROD_REQUEST_DATE, P400.PROD_REQUEST_QTY, P501.INS_DATE AS PROD_DATETIME, ZTBINSPECTOUTPUTTB.OUTPUT_DATETIME, ZTBINSPECTOUTPUTTB.OUTPUT_QTY_EA, ZTBINSPECTOUTPUTTB.REMARK FROM ZTBINSPECTOUTPUTTB JOIN P501 ON(ZTBINSPECTOUTPUTTB.PROCESS_LOT_NO = P501.PROCESS_LOT_NO) JOIN(SELECT DISTINCT PROD_REQUEST_NO, PROCESS_IN_DATE, PROCESS_IN_NO FROM P500) AS P500_A ON(P500_A.PROCESS_IN_DATE = P501.PROCESS_IN_DATE AND P500_A.PROCESS_IN_NO = P501.PROCESS_IN_NO) JOIN P400 ON(P500_A.PROD_REQUEST_NO = P400.PROD_REQUEST_NO) JOIN M110 ON(M110.CUST_CD = P400.CUST_CD) JOIN M100 ON(M100.G_CODE = P400.G_CODE) JOIN M010 ON(M010.EMPL_NO = ZTBINSPECTOUTPUTTB.EMPL_NO)

            //string strQuery = "SELECT ZTBINSPECTOUTPUTTB.INSPECT_OUTPUT_ID,M110.CUST_NAME_KD, M010.EMPL_NAME,ZTBINSPECTOUTPUTTB.G_CODE,M100.G_NAME,M100.PROD_TYPE,M100.G_NAME_KD,ZTBINSPECTOUTPUTTB.PROD_REQUEST_NO,P400.PROD_REQUEST_DATE,P400.PROD_REQUEST_QTY,ZTBINSPECTOUTPUTTB.PROCESS_LOT_NO,P501_A.INS_DATE AS PROD_DATETIME, ZTBINSPECTOUTPUTTB.OUTPUT_DATETIME,ZTBINSPECTOUTPUTTB.OUTPUT_QTY_EA,ZTBINSPECTOUTPUTTB.REMARK,P400.EMPL_NO AS PIC_KD,CASE  WHEN (DATEPART(HOUR,OUTPUT_DATETIME) >=8 AND DATEPART(HOUR,OUTPUT_DATETIME) <20) THEN 'CA NGAY'  ELSE 'CA DEM' END AS CA_LAM_VIEC,  CASE  WHEN DATEPART(HOUR,OUTPUT_DATETIME) < 8  THEN CONVERT(date,DATEADD(DAY,-1,OUTPUT_DATETIME))  ELSE CONVERT(date,OUTPUT_DATETIME) END  AS NGAY_LAM_VIEC  FROM ZTBINSPECTOUTPUTTB  LEFT JOIN M010 ON (M010.EMPL_NO = ZTBINSPECTOUTPUTTB.EMPL_NO)  LEFT JOIN P400 ON (P400.PROD_REQUEST_NO = ZTBINSPECTOUTPUTTB.PROD_REQUEST_NO)  LEFT JOIN (SELECT * FROM P501 WHERE INS_DATE>'2021-07-01') AS P501_A ON (P501_A.PROCESS_LOT_NO = ZTBINSPECTOUTPUTTB.PROCESS_LOT_NO)  LEFT JOIN M100 ON (M100.G_CODE = ZTBINSPECTOUTPUTTB.G_CODE)  LEFT JOIN M110 ON (M110.CUST_CD = P400.CUST_CD) " + condition + " ORDER BY INSPECT_OUTPUT_ID DESC";
            string strQuery = "SELECT ZTBINSPECTOUTPUTTB.INSPECT_OUTPUT_ID,M110.CUST_NAME_KD, M010.EMPL_NAME,ZTBINSPECTOUTPUTTB.G_CODE,M100.G_NAME,M100.PROD_TYPE,M100.G_NAME_KD,ZTBINSPECTOUTPUTTB.PROD_REQUEST_NO,P400.PROD_REQUEST_DATE,P400.PROD_REQUEST_QTY,ZTBINSPECTOUTPUTTB.PROCESS_LOT_NO, ZTBINSPECTOUTPUTTB.OUTPUT_DATETIME,ZTBINSPECTOUTPUTTB.OUTPUT_QTY_EA,ZTBINSPECTOUTPUTTB.REMARK,P400.EMPL_NO AS PIC_KD,CASE  WHEN (DATEPART(HOUR,OUTPUT_DATETIME) >=8 AND DATEPART(HOUR,OUTPUT_DATETIME) <20) THEN 'CA NGAY'  ELSE 'CA DEM' END AS CA_LAM_VIEC,  CASE  WHEN DATEPART(HOUR,OUTPUT_DATETIME) < 8  THEN CONVERT(date,DATEADD(DAY,-1,OUTPUT_DATETIME))  ELSE CONVERT(date,OUTPUT_DATETIME) END  AS NGAY_LAM_VIEC  FROM ZTBINSPECTOUTPUTTB  LEFT JOIN M010 ON (M010.EMPL_NO = ZTBINSPECTOUTPUTTB.EMPL_NO)  LEFT JOIN P400 ON (P400.PROD_REQUEST_NO = ZTBINSPECTOUTPUTTB.PROD_REQUEST_NO)  LEFT JOIN M100 ON (M100.G_CODE = ZTBINSPECTOUTPUTTB.G_CODE)  LEFT JOIN M110 ON (M110.CUST_CD = P400.CUST_CD) " + condition + " ORDER BY INSPECT_OUTPUT_ID DESC";

            //MessageBox.Show(strQuery);
            result = config.GetData(strQuery);
            return result;

        }


        public DataTable report_inspection_all_input_data(string condition)
        {
            
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT ZTBINSPECTINPUTTB.INSPECT_INPUT_ID,M110.CUST_NAME_KD, M010.EMPL_NAME,ZTBINSPECTINPUTTB.G_CODE,M100.G_NAME,M100.PROD_TYPE,M100.G_NAME_KD,ZTBINSPECTINPUTTB.PROD_REQUEST_NO,P400.PROD_REQUEST_DATE,P400.PROD_REQUEST_QTY,ZTBINSPECTINPUTTB.PROCESS_LOT_NO,P501_A.INS_DATE AS PROD_DATETIME, ZTBINSPECTINPUTTB.INPUT_DATETIME,ZTBINSPECTINPUTTB.INPUT_QTY_EA,ZTBINSPECTINPUTTB.INPUT_QTY_KG,ZTBINSPECTINPUTTB.REMARK,ZTBINSPECTINPUTTB.CNDB_ENCODES,P400.EMPL_NO AS PIC_KD  FROM ZTBINSPECTINPUTTB  LEFT JOIN M010 ON (M010.EMPL_NO = ZTBINSPECTINPUTTB.EMPL_NO)  LEFT JOIN P400 ON (P400.PROD_REQUEST_NO = ZTBINSPECTINPUTTB.PROD_REQUEST_NO)  LEFT JOIN (SELECT * FROM P501 WHERE INS_DATE>'2021-07-01') AS P501_A ON (P501_A.PROCESS_LOT_NO = ZTBINSPECTINPUTTB.PROCESS_LOT_NO)  LEFT JOIN M100 ON (M100.G_CODE = ZTBINSPECTINPUTTB.G_CODE)  LEFT JOIN M110 ON (M110.CUST_CD = P400.CUST_CD) " + condition + " ORDER BY INSPECT_INPUT_ID DESC ";
            result = config.GetData(strQuery);
            return result;

        }

      // inspection part start


        public string getPODDate(string G_CODE, string PO_NO)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT PO_DATE FROM ZTBPOTable WHERE G_CODE='" + G_CODE +  "' AND PO_NO='" + PO_NO + "'";
            result = config.GetData(strQuery);
            string output ="";
            if (result.Rows.Count > 0)
            {
                foreach (DataRow row in result.Rows)
                {
                    output = row[0].ToString();
                }
            }
            else
            {
                output = "0";
            }
            return output;
        }

        public int getVer()
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT * FROM ZBTVERTABLE";
            result = config.GetData(strQuery);
            int output = 0;
            if(result.Rows.Count>0)
            {
                foreach (DataRow row in result.Rows)
                {
                    output = int.Parse(row[0].ToString());
                }
            }
            else
            {
                output = 0;
            }
            return output;
        }

        public DataTable readHistory()
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT HISTORY_ID, ZTBHistory.EMPL_NO, M010.EMPL_NAME, TABLENAME, ACTIONS, CONTENT, ID, EDITTIME FROM ZTBHistory JOIN M010 ON (M010.EMPL_NO = ZTBHistory.EMPL_NO) ORDER BY EDITTIME DESC";
            result = config.GetData(strQuery);
            return result;
        }

        public DataTable writeHistory(string CTR_CD, string EMPL_NO, string TABLENAME, string ACTIONS, string CONTENT, string ID)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "INSERT INTO ZTBHistory (CTR_CD,EMPL_NO,TABLENAME,ACTIONS,CONTENT,ID) VALUES('002','"+ EMPL_NO + "','"+ TABLENAME + "','"+ ACTIONS +"','"+ CONTENT + "',"+ ID + ")";
            result = config.GetData(strQuery);
            return result;
        }


        public DataTable info_getCODEBOM(string g_code)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT M100.G_NAME, M100.G_CODE, M140.M_CODE, M090.M_NAME, M140.M_QTY, M090.WIDTH_CD , M100.ROLE_EA_QTY AS PACKING_QTY, G_WIDTH,G_LENGTH, (G_LENGTH + G_LG) AS PD, G_C AS CAVITY,M090.INS_EMPL, M090.UPD_EMPL FROM M140 JOIN M090 ON M140.M_CODE = M090.M_CODE JOIN M100 ON M100.G_CODE = M140.G_CODE WHERE M100.G_CODE ='" + g_code + "'";
            result = config.GetData(strQuery);
            return result;
        }

        public DataTable info_getCODEInfo(string g_name)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT M100.G_CODE, M100.G_NAME, M100.G_NAME_KD, M110.CUST_NAME_KD, M110.CUST_NAME, M100.ROLE_EA_QTY AS PACKING_QTY, G_WIDTH,G_LENGTH, (G_LENGTH + G_LG) AS PD, G_C AS CAVITY, M100.PROD_PROJECT, M100.PROD_MODEL, M100.PROD_TYPE, M100.PROD_MAIN_MATERIAL, M100.USE_YN  FROM M100 JOIN M110 ON M100.CUST_CD = M110.CUST_CD WHERE G_NAME LIKE '%" + g_name + "%'";
            result = config.GetData(strQuery);
            return result;
        }


        public int getlastFCSTWeekNum(int fcstyear)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT ISNULL(MAX(FCSTWEEKNO),-1) AS MAXFCSTWEEKNO FROM ZTBFCSTTB WHERE FCSTYEAR=" + fcstyear ;
            result = config.GetData(strQuery);
            int maxweeknum = 0;            
            if(result.Rows.Count >0)
            {
                maxweeknum =  int.Parse(result.Rows[0]["MAXFCSTWEEKNO"].ToString());
            }
            else
            {
                maxweeknum = -1;
            }
            return maxweeknum;
        }



        public DataTable report_getEmployeeList()
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT EMPL_NO, EMPL_NAME, M060.DEPT_NAME FROM M010 JOIN M060 ON M010.DEPT_CD = M060.DEPT_CD";
            result = config.GetData(strQuery);
            return result;

        }


        public DataTable report_getCustomerList()
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT CUST_CD, CUST_NAME, CUST_NAME_KD FROM M110";
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable report_getCustomerList1(string searchValue)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT CUST_CD, CUST_NAME, CUST_NAME_KD FROM M110 WHERE " + searchValue;
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable report_addCustomer(string addValue)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "INSERT INTO M110 (CTR_CD,CUST_CD, CUST_NAME, CUST_NAME_KD) VALUES " + addValue;
            result = config.GetData(strQuery);
            return result;

        }
        public DataTable report_updateCustomer(string updateValue)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "UPDATE M110 " + updateValue;
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable report_MaterialList()
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT M_CODE, M_NAME , WIDTH_CD,CONCAT(M_NAME, '|',WIDTH_CD ,'|', M_CODE) AS M_NAME_SIZEZ FROM M090";
            result = config.GetData(strQuery);
            return result;

        }



        public DataTable report_getCodeList()
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT M100.G_CODE, M100.G_NAME, M110.CUST_NAME_KD, M110.CUST_NAME, M100.PROD_LAST_PRICE, M100.ROLE_EA_QTY AS PACKING_QTY, G_WIDTH,G_LENGTH, (G_LENGTH + G_LG) AS PD, G_C AS CAVITY, M100.PROD_PROJECT, M100.PROD_MODEL, M100.PROD_TYPE, M100.PROD_MAIN_MATERIAL  FROM M100 JOIN M110 ON M100.CUST_CD = M110.CUST_CD";
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable report_NoData()
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT DISTINCT M100.G_CODE, M100.G_NAME, M100.G_NAME_KD, M100.PROD_TYPE, M100.PROD_MAIN_MATERIAL, M100.PROD_MODEL, M100.PROD_PROJECT FROM M100 WHERE G_CODE IN (SELECT DISTINCT G_CODE FROM ZTBINSPECTINPUTTB)  OR G_CODE IN (SELECT DISTINCT G_CODE FROM ZTBPOTable)  OR G_CODE IN (SELECT DISTINCT G_CODE FROM ZTBFCSTTB)  AND (PROD_TYPE is null OR  PROD_MODEL is null OR PROD_MAIN_MATERIAL is null OR PROD_PROJECT is null OR PROD_TYPE = '' OR  PROD_MODEL = '' OR PROD_MAIN_MATERIAL = '' OR PROD_PROJECT = '' OR G_NAME_KD is null OR G_NAME_KD ='')";
            result = config.GetData(strQuery);
            return result;
        }
        public DataTable report_QLSX()
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT G_CODE, G_NAME, G_NAME_KD, PROD_TYPE, PROD_MODEL, PROD_PROJECT, PROD_MAIN_MATERIAL, FACTORY, EQ1, EQ2, Setting1, Setting2, UPH1, UPH2, Step1, Step2, NOTE FROM M100 WHERE ( 	G_CODE IN (SELECT DISTINCT AA.G_CODE FROM 	(SELECT ZTBPOTable.G_CODE, SUM(ZTBPOTable.PO_QTY) AS TOTAL_PO_QTY, SUM(isnull(ZTBDelivery_A.DELIVERED_QTY,0)) AS TOTAL_DELIVERED_QTY, (SUM(ZTBPOTable.PO_QTY) -  SUM(isnull(ZTBDelivery_A.DELIVERED_QTY,0))) AS PO_BALANCE   FROM ZTBPOTable 	LEFT JOIN (SELECT G_CODE, CUST_CD, PO_NO, SUM(DELIVERY_QTY) AS DELIVERED_QTY FROM ZTBDelivery GROUP BY G_CODE, CUST_CD, PO_NO) AS ZTBDelivery_A ON (ZTBPOTable.G_CODE = ZTBDelivery_A.G_CODE AND ZTBPOTable.CUST_CD = ZTBDelivery_A.CUST_CD AND ZTBPOTable.PO_NO = ZTBDelivery_A.PO_NO) 	GROUP BY ZTBPOTable.G_CODE) AS AA 	WHERE AA.PO_BALANCE <> 0 	) 	OR 	(G_CODE IN (SELECT DISTINCT G_CODE FROM ZTBPLANTB WHERE PLAN_DATE BETWEEN  DATEADD(DAY,-14, GETDATE()) AND GETDATE()) OR G_CODE IN  (SELECT DISTINCT G_CODE FROM ZTBINSPECTINPUTTB WHERE INPUT_DATETIME between  DATEADD(DAY,-14, GETDATE()) AND  DATEADD(DAY,1, GETDATE()))) )AND (M100.FACTORY is null or M100.EQ1 is null)";
            result = config.GetData(strQuery);
            return result;
        }
        public DataTable report_QLSX_validating()
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT G_CODE, G_NAME, G_NAME_KD, PROD_TYPE, PROD_MODEL, PROD_PROJECT, PROD_MAIN_MATERIAL, FACTORY, EQ1, EQ2, Setting1, Setting2, UPH1, UPH2, Step1, Step2, NOTE FROM M100 WHERE (FACTORY is not null  AND FACTORY <>'' AND EQ1 <> 'VD') AND (	(EQ1 is not null AND  (Setting1 =0 OR UPH1 = 0 OR Step1 =0)) 	 OR 	(EQ2 <> '' AND  (Setting2 =0 OR UPH2 = 0 OR Step2 =0)) )";
            result = config.GetData(strQuery);
            return result;
        }

        public DataTable traBOMCAPA(string condition)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = $"SELECT G_CODE, G_NAME, G_NAME_KD, PROD_TYPE, PROD_MODEL, PROD_PROJECT, PROD_MAIN_MATERIAL, FACTORY, EQ1, EQ2, Setting1, Setting2, UPH1, UPH2, Step1, Step2, NOTE FROM M100 " + condition;
            result = config.GetData(strQuery);
            return result;
        }

        public DataTable traLICHSUINPUTLIEU(string condition)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = $"SELECT P500.PROCESS_IN_DATE, P500.PROCESS_IN_NO, P500.G_CODE, P500.PROD_REQUEST_NO, P500.PROD_REQUEST_DATE, M100.G_NAME, M090.M_NAME, M090.WIDTH_CD, P500.M_LOT_NO, M010.EMPL_NAME, P500.EQUIPMENT_CD, P500.INS_DATE FROM P500 JOIN P400 ON (P500.PROD_REQUEST_NO = P400.PROD_REQUEST_NO AND P500.PROD_REQUEST_DATE = P400.PROD_REQUEST_DATE) JOIN M100 ON (P500.G_CODE = M100.G_CODE) JOIN M090 ON (M090.M_CODE = P500.M_CODE) JOIN M010 ON (M010.EMPL_NO = P500.EMPL_NO) {condition} ORDER BY P500.INS_DATE DESC";
            result = config.GetData(strQuery);
            return result;
        }

        public DataTable report_checkBTP(string condition)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT ZTB_HALF_GOODS.HG_ID, ZTB_HALF_GOODS.FACTORY, ZTB_HALF_GOODS.UPDATE_DATE, ZTB_HALF_GOODS.PROD_REQUEST_NO, ZTB_HALF_GOODS.G_CODE, M100.G_NAME, M100.G_NAME_KD,  ZTB_HALF_GOODS.ROLL_QTY, ZTB_HALF_GOODS.BTP_POSITION, ZTB_HALF_GOODS.BTP_MET, ZTB_HALF_GOODS.BTP_PD, ZTB_HALF_GOODS.BTP_CAVITY, ZTB_HALF_GOODS.BTP_QTY_EA, ZTB_HALF_GOODS.INS_EMPL_NO, ZTB_HALF_GOODS.INS_DATETIME FROM ZTB_HALF_GOODS LEFT JOIN M100 ON ( M100.G_CODE = ZTB_HALF_GOODS.G_CODE) " + condition;
            result = config.GetData(strQuery);
            return result;
        }

        public DataTable report_checkUP_BTP_today()
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT * FROM ZTB_HALF_GOODS WHERE UPDATE_DATE = CONVERT(date,GETDATE())";
            result = config.GetData(strQuery);
            return result;
        }

        public DataTable report_checkUP_TONKIEM_today()
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT * FROM ZTB_WAIT_INSPECT WHERE UPDATE_DATE = CONVERT(date,GETDATE()) AND CALAMVIEC='DEM'";
            result = config.GetData(strQuery);
            return result;
        }


        public DataTable report_checkBTP2(string condition)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT ZTB_HALF_GOODS.G_CODE, M100.G_NAME, SUM(BTP_QTY_EA) AS BTP_QTY_EA FROM ZTB_HALF_GOODS JOIN M100 ON (M100.G_CODE = ZTB_HALF_GOODS.G_CODE) " + condition + " GROUP BY ZTB_HALF_GOODS.G_CODE, M100.G_NAME";
            //MessageBox.Show(strQuery);
            result = config.GetData(strQuery);
            return result;
        }

        public DataTable report_TONKHOFULL(string condition)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = $"SELECT M100.G_CODE, M100.G_NAME, M100.G_NAME_KD, isnull(TONKIEM.INSPECT_BALANCE_QTY,0) AS CHO_KIEM, isnull(TONKIEM.WAIT_CS_QTY,0) AS CHO_CS_CHECK,isnull(TONKIEM.WAIT_SORTING_RMA,0) CHO_KIEM_RMA, isnull(TONKIEM.TOTAL_WAIT,0) AS TONG_TON_KIEM, isnull(BTP.BTP_QTY_EA,0) AS BTP, isnull(THANHPHAM.TONKHO,0) AS TON_TP, isnull(tbl_Block_table2.Block_Qty,0) AS BLOCK_QTY, (isnull(TONKIEM.TOTAL_WAIT,0) + isnull(BTP.BTP_QTY_EA,0)+ isnull(THANHPHAM.TONKHO,0) - isnull(tbl_Block_table2.Block_Qty,0)) AS GRAND_TOTAL_STOCK FROM M100 LEFT JOIN ( SELECT Product_MaVach, isnull([IN],0) AS NHAPKHO, isnull([OUT],0) AS XUATKHO, (isnull([IN],0)- isnull([OUT],0)) AS TONKHO FROM ( SELECT Product_Mavach, IO_Type, IO_Qty FROM tbl_InputOutput ) AS SourceTable PIVOT ( SUM(IO_Qty) FOR IO_Type IN ([IN], [OUT]) ) AS PivotTable ) AS THANHPHAM ON (THANHPHAM.Product_MaVach = M100.G_CODE) LEFT JOIN ( SELECT ZTB_WAIT_INSPECT.G_CODE, M100.G_NAME, M100.G_NAME_KD, SUM(INSPECT_BALANCE_QTY) AS INSPECT_BALANCE_QTY, SUM(WAIT_CS_QTY) AS WAIT_CS_QTY, SUM(WAIT_SORTING_RMA) AS WAIT_SORTING_RMA, SUM(INSPECT_BALANCE_QTY+ WAIT_CS_QTY+ WAIT_SORTING_RMA) AS TOTAL_WAIT FROM ZTB_WAIT_INSPECT JOIN M100 ON ( M100.G_CODE = ZTB_WAIT_INSPECT.G_CODE) WHERE UPDATE_DATE=CONVERT(date,GETDATE()) AND CALAMVIEC = 'DEM' GROUP BY ZTB_WAIT_INSPECT.G_CODE, M100.G_NAME, M100.G_NAME_KD) AS TONKIEM ON (THANHPHAM.Product_MaVach = TONKIEM.G_CODE) LEFT JOIN ( SELECT Product_MaVach, SUM(Block_Qty) AS Block_Qty from tbl_Block2 GROUP BY Product_MaVach ) AS tbl_Block_table2 ON (tbl_Block_table2.Product_MaVach= M100.G_CODE) LEFT JOIN ( SELECT ZTB_HALF_GOODS.G_CODE, M100.G_NAME, SUM(BTP_QTY_EA) AS BTP_QTY_EA FROM ZTB_HALF_GOODS JOIN M100 ON (M100.G_CODE = ZTB_HALF_GOODS.G_CODE) WHERE UPDATE_DATE = CONVERT(date,GETDATE()) GROUP BY ZTB_HALF_GOODS.G_CODE, M100.G_NAME) AS BTP ON (BTP.G_CODE = THANHPHAM.Product_MaVach) {condition} ";
            //MessageBox.Show(strQuery);
            result = config.GetData(strQuery);
            return result;
        }

        public DataTable report_TONKHOFULLKD(string condition)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = $"SELECT M100.G_NAME_KD, SUM(isnull(TONKIEM.INSPECT_BALANCE_QTY,0)) AS CHO_KIEM, SUM(isnull(TONKIEM.WAIT_CS_QTY,0)) AS CHO_CS_CHECK,SUM(isnull(TONKIEM.WAIT_SORTING_RMA,0)) AS CHO_KIEM_RMA, SUM(isnull(TONKIEM.TOTAL_WAIT,0)) AS TONG_TON_KIEM, SUM(isnull(BTP.BTP_QTY_EA,0)) AS BTP, SUM(isnull(THANHPHAM.TONKHO,0)) AS TON_TP, SUM(isnull(tbl_Block_table2.Block_Qty,0)) AS BLOCK_QTY, SUM((isnull(TONKIEM.TOTAL_WAIT,0)) + isnull(BTP.BTP_QTY_EA,0)+ isnull(THANHPHAM.TONKHO,0) - isnull(tbl_Block_table2.Block_Qty,0)) AS GRAND_TOTAL_STOCK FROM M100 LEFT JOIN ( SELECT Product_MaVach, isnull([IN],0) AS NHAPKHO, isnull([OUT],0) AS XUATKHO, (isnull([IN],0)- isnull([OUT],0)) AS TONKHO FROM ( SELECT Product_Mavach, IO_Type, IO_Qty FROM tbl_InputOutput ) AS SourceTable PIVOT ( SUM(IO_Qty) FOR IO_Type IN ([IN], [OUT]) ) AS PivotTable ) AS THANHPHAM ON (THANHPHAM.Product_MaVach = M100.G_CODE) LEFT JOIN ( SELECT ZTB_WAIT_INSPECT.G_CODE, M100.G_NAME, M100.G_NAME_KD, SUM(INSPECT_BALANCE_QTY) AS INSPECT_BALANCE_QTY, SUM(WAIT_CS_QTY) AS WAIT_CS_QTY, SUM(WAIT_SORTING_RMA) AS WAIT_SORTING_RMA, SUM(INSPECT_BALANCE_QTY+ WAIT_CS_QTY+ WAIT_SORTING_RMA) AS TOTAL_WAIT FROM ZTB_WAIT_INSPECT JOIN M100 ON ( M100.G_CODE = ZTB_WAIT_INSPECT.G_CODE) WHERE UPDATE_DATE=CONVERT(date,GETDATE()) AND CALAMVIEC = 'DEM' GROUP BY ZTB_WAIT_INSPECT.G_CODE, M100.G_NAME, M100.G_NAME_KD) AS TONKIEM ON (THANHPHAM.Product_MaVach = TONKIEM.G_CODE) LEFT JOIN ( SELECT Product_MaVach, SUM(Block_Qty) AS Block_Qty from tbl_Block2 GROUP BY Product_MaVach ) AS tbl_Block_table2 ON (tbl_Block_table2.Product_MaVach= M100.G_CODE) LEFT JOIN ( SELECT ZTB_HALF_GOODS.G_CODE, M100.G_NAME, SUM(BTP_QTY_EA) AS BTP_QTY_EA FROM ZTB_HALF_GOODS JOIN M100 ON (M100.G_CODE = ZTB_HALF_GOODS.G_CODE) WHERE UPDATE_DATE = CONVERT(date,GETDATE()) GROUP BY ZTB_HALF_GOODS.G_CODE, M100.G_NAME) AS BTP ON (BTP.G_CODE = THANHPHAM.Product_MaVach) {condition} GROUP BY M100.G_NAME_KD";
            //MessageBox.Show(strQuery);
            result = config.GetData(strQuery);
            return result;
        }


        public DataTable report_TONKHOTACH(string condition)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = $"SELECT isnull(THANHPHAM.WH_Name,'NO_STOCK') AS KHO_NAME, tbl_Location.LC_Name, M100.G_CODE, M100.G_NAME, M100.G_NAME_KD, isnull(THANHPHAM.NHAPKHO,0) AS NHAPKHO, isnull(THANHPHAM.XUATKHO,0) AS XUATKHO, isnull(THANHPHAM.TONKHO,0) AS TONKHO, isnull(tbl_Block_table2.Block_Qty,0) AS BLOCK_QTY, ( isnull(THANHPHAM.TONKHO,0)-isnull(tbl_Block_table2.Block_Qty,0)) AS GRAND_TOTAL_TP FROM M100 LEFT JOIN ( SELECT Product_MaVach, WH_Name, isnull([IN],0) AS NHAPKHO, isnull([OUT],0) AS XUATKHO, (isnull([IN],0)- isnull([OUT],0)) AS TONKHO FROM ( SELECT Product_Mavach, WH_Name, IO_Type, IO_Qty FROM tbl_InputOutput ) AS SourceTable PIVOT ( SUM(IO_Qty) FOR IO_Type IN ([IN], [OUT]) ) AS PivotTable ) AS THANHPHAM ON (THANHPHAM.Product_MaVach = M100.G_CODE) LEFT JOIN ( SELECT Product_MaVach, WH_Name, SUM(Block_Qty) AS Block_Qty from tbl_Block2 GROUP BY Product_MaVach,WH_Name ) AS tbl_Block_table2 ON (tbl_Block_table2.Product_MaVach= THANHPHAM.Product_MaVach AND tbl_Block_table2.WH_Name= THANHPHAM.WH_Name) LEFT JOIN tbl_Location ON (tbl_Location.Product_MaVach = THANHPHAM.Product_MaVach AND tbl_Location.WH_Name = THANHPHAM.WH_Name) {condition} ";
            //MessageBox.Show(strQuery);
            result = config.GetData(strQuery);
            return result;
        }


        public DataTable report_TONKHO_INPUT(string condition)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = $"SELECT tbl_InputOutput.Product_MaVach, M100.G_NAME, M100.G_NAME_KD, tbl_InputOutput.Customer_ShortName, tbl_InputOutput.IO_Date, CONVERT(datetime,tbl_InputOutput.IO_Time) AS INPUT_DATETIME, tbl_InputOutput.IO_Shift ,tbl_InputOutput.IO_Type, tbl_InputOutput.IO_Qty FROM tbl_InputOutput LEFT JOIN M100 ON (M100.G_CODE= tbl_InputOutput.Product_MaVach) {condition} ORDER BY tbl_InputOutput.IO_Time DESC ";
            //MessageBox.Show(strQuery);
            result = config.GetData(strQuery);
            return result;
        }

        public DataTable report_TONKHO_OUTPUT(string condition)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = $"SELECT tbl_InputOutput.Product_MaVach, M100.G_NAME, M100.G_NAME_KD, tbl_InputOutput.Customer_ShortName, tbl_InputOutput.IO_Date, CONVERT(datetime,tbl_InputOutput.IO_Time) AS OUTPUT_DATETIME, tbl_InputOutput.IO_Shift ,tbl_InputOutput.IO_Type, tbl_InputOutput.IO_Qty FROM tbl_InputOutput LEFT JOIN M100 ON (M100.G_CODE= tbl_InputOutput.Product_MaVach) {condition} ORDER BY tbl_InputOutput.IO_Time DESC ";
            //MessageBox.Show(strQuery);
            result = config.GetData(strQuery);
            return result;
        }




        public DataTable report_checkCK(string condition)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT  ZTB_WAIT_INSPECT.WI_ID, ZTB_WAIT_INSPECT.FACTORY, ZTB_WAIT_INSPECT.UPDATE_DATE, ZTB_WAIT_INSPECT.PROD_REQUEST_NO, ZTB_WAIT_INSPECT.G_CODE, M100.G_NAME, M100.G_NAME_KD, M100.PROD_TYPE, ZTB_WAIT_INSPECT.INSPECT_BALANCE_QTY, ZTB_WAIT_INSPECT.WAIT_CS_QTY, ZTB_WAIT_INSPECT.WAIT_SORTING_RMA, (ZTB_WAIT_INSPECT.INSPECT_BALANCE_QTY+ ZTB_WAIT_INSPECT.WAIT_CS_QTY+ ZTB_WAIT_INSPECT.WAIT_SORTING_RMA) AS TOTAL_WAIT, ZTB_WAIT_INSPECT.INS_EMPL_NO, ZTB_WAIT_INSPECT.INS_DATETIME, ZTB_WAIT_INSPECT.REMARK, ZTB_WAIT_INSPECT.CALAMVIEC FROM ZTB_WAIT_INSPECT LEFT JOIN M100 ON(M100.G_CODE = ZTB_WAIT_INSPECT.G_CODE) " + condition;
            result = config.GetData(strQuery);
            return result;
        }

        public DataTable report_checkCK2(string condition)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT ZTB_WAIT_INSPECT.G_CODE, M100.G_NAME, M100.G_NAME_KD, SUM(INSPECT_BALANCE_QTY) AS INSPECT_BALANCE_QTY, SUM(WAIT_CS_QTY) AS WAIT_CS_QTY, SUM(WAIT_SORTING_RMA) AS WAIT_SORTING_RMA,  SUM(INSPECT_BALANCE_QTY+ WAIT_CS_QTY+ WAIT_SORTING_RMA) AS TOTAL_WAIT FROM ZTB_WAIT_INSPECT JOIN M100 ON ( M100.G_CODE = ZTB_WAIT_INSPECT.G_CODE) " + condition + " GROUP BY ZTB_WAIT_INSPECT.G_CODE, M100.G_NAME, M100.G_NAME_KD";
            result = config.GetData(strQuery);
            return result;
        }


        public DataTable report_OverDueByTYPE()
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT DISTINCT XX.PROD_TYPE, COUNT(OVERDUE) AS OVERDUE FROM (SELECT AA.PO_NO,  M100.PROD_TYPE, M100.PROD_MAIN_MATERIAL, ZTBPOTable.PO_DATE,ZTBPOTable.RD_DATE, M110.CUST_NAME_KD, M100.G_NAME, M010.EMPL_NAME, AA.G_CODE, ZTBPOTable.PO_QTY, ZTBPOTable.PROD_PRICE, AA.TotalDelivered as TOTAL_DELIVERED, (ZTBPOTable.PO_QTY-AA.TotalDelivered) As PO_BALANCE,(ZTBPOTable.PO_QTY*ZTBPOTable.PROD_PRICE) As PO_AMOUNT , (AA.TotalDelivered*ZTBPOTable.PROD_PRICE) As DELIVERED_AMOUNT, ((ZTBPOTable.PO_QTY-AA.TotalDelivered)*ZTBPOTable.PROD_PRICE) As BALANCE_AMOUNT,DATEPART( MONTH, PO_DATE) AS POMONTH, DATEPART( ISOWK, PO_DATE) AS POWEEKNUM, YEAR(PO_DATE) As PO_YEAR, CASE     WHEN (ZTBPOTable.RD_DATE < GETDATE()-1) AND ((ZTBPOTable.PO_QTY-AA.TotalDelivered) <>0) THEN 'OVER'        ELSE 'OK' END AS OVERDUE, ZTBPOTable.REMARK, ZTBPOTable.PO_ID FROM (SELECT ZTBPOTable.EMPL_NO, ZTBPOTable.CUST_CD, ZTBPOTable.G_CODE, ZTBPOTable.PO_NO, isnull(SUM(ZTBDelivery.DELIVERY_QTY),0) AS TotalDelivered FROM ZTBPOTable  LEFT JOIN ZTBDelivery ON (ZTBDelivery.CTR_CD = ZTBPOTable.CTR_CD AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) GROUP BY ZTBPOTable.CTR_CD,ZTBPOTable.EMPL_NO,ZTBPOTable.G_CODE,ZTBPOTable.CUST_CD,ZTBPOTable.PO_NO) AS AA LEFT JOIN M010 ON (M010.EMPL_NO = AA.EMPL_NO) LEFT JOIN M100 ON (M100.G_CODE = AA.G_CODE) LEFT JOIN ZTBPOTable ON (AA.CUST_CD = ZTBPOTable.CUST_CD AND AA.G_CODE = ZTBPOTable.G_CODE AND AA.PO_NO = ZTBPOTable.PO_NO) JOIN M110 ON (M110.CUST_CD = AA.CUST_CD) ) AS XX WHERE XX.OVERDUE= 'OVER' GROUP BY XX.PROD_TYPE ORDER BY OVERDUE DESC";
            result = config.GetData(strQuery);
            return result;
        }

        public DataTable report_OverDueByCustomer()
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT DISTINCT XX.CUST_NAME_KD, COUNT(OVERDUE) AS OVERDUE FROM (SELECT AA.PO_NO,  M100.PROD_TYPE, M100.PROD_MAIN_MATERIAL, ZTBPOTable.PO_DATE,ZTBPOTable.RD_DATE, M110.CUST_NAME_KD, M100.G_NAME, M010.EMPL_NAME, AA.G_CODE, ZTBPOTable.PO_QTY, ZTBPOTable.PROD_PRICE, AA.TotalDelivered as TOTAL_DELIVERED, (ZTBPOTable.PO_QTY-AA.TotalDelivered) As PO_BALANCE,(ZTBPOTable.PO_QTY*ZTBPOTable.PROD_PRICE) As PO_AMOUNT , (AA.TotalDelivered*ZTBPOTable.PROD_PRICE) As DELIVERED_AMOUNT, ((ZTBPOTable.PO_QTY-AA.TotalDelivered)*ZTBPOTable.PROD_PRICE) As BALANCE_AMOUNT,DATEPART( MONTH, PO_DATE) AS POMONTH, DATEPART( ISOWK, PO_DATE) AS POWEEKNUM, YEAR(PO_DATE) As PO_YEAR, CASE     WHEN (ZTBPOTable.RD_DATE < GETDATE()-1) AND ((ZTBPOTable.PO_QTY-AA.TotalDelivered) <>0) THEN 'OVER'        ELSE 'OK' END AS OVERDUE, ZTBPOTable.REMARK, ZTBPOTable.PO_ID FROM (SELECT ZTBPOTable.EMPL_NO, ZTBPOTable.CUST_CD, ZTBPOTable.G_CODE, ZTBPOTable.PO_NO, isnull(SUM(ZTBDelivery.DELIVERY_QTY),0) AS TotalDelivered FROM ZTBPOTable  LEFT JOIN ZTBDelivery ON (ZTBDelivery.CTR_CD = ZTBPOTable.CTR_CD AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) GROUP BY ZTBPOTable.CTR_CD,ZTBPOTable.EMPL_NO,ZTBPOTable.G_CODE,ZTBPOTable.CUST_CD,ZTBPOTable.PO_NO) AS AA LEFT JOIN M010 ON (M010.EMPL_NO = AA.EMPL_NO) LEFT JOIN M100 ON (M100.G_CODE = AA.G_CODE) LEFT JOIN ZTBPOTable ON (AA.CUST_CD = ZTBPOTable.CUST_CD AND AA.G_CODE = ZTBPOTable.G_CODE AND AA.PO_NO = ZTBPOTable.PO_NO) JOIN M110 ON (M110.CUST_CD = AA.CUST_CD) ) AS XX WHERE XX.OVERDUE= 'OVER' GROUP BY XX.CUST_NAME_KD ORDER BY OVERDUE DESC";
            result = config.GetData(strQuery);
            return result;
        }
        
        public DataTable report_OverDueByPIC()
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT DISTINCT XX.EMPL_NAME, COUNT(OVERDUE) AS OVERDUE FROM (SELECT AA.PO_NO,  M100.PROD_TYPE, M100.PROD_MAIN_MATERIAL, ZTBPOTable.PO_DATE,ZTBPOTable.RD_DATE, M110.CUST_NAME_KD, M100.G_NAME, M010.EMPL_NAME, AA.G_CODE, ZTBPOTable.PO_QTY, ZTBPOTable.PROD_PRICE, AA.TotalDelivered as TOTAL_DELIVERED, (ZTBPOTable.PO_QTY-AA.TotalDelivered) As PO_BALANCE,(ZTBPOTable.PO_QTY*ZTBPOTable.PROD_PRICE) As PO_AMOUNT , (AA.TotalDelivered*ZTBPOTable.PROD_PRICE) As DELIVERED_AMOUNT, ((ZTBPOTable.PO_QTY-AA.TotalDelivered)*ZTBPOTable.PROD_PRICE) As BALANCE_AMOUNT,DATEPART( MONTH, PO_DATE) AS POMONTH, DATEPART( ISOWK, PO_DATE) AS POWEEKNUM, YEAR(PO_DATE) As PO_YEAR, CASE     WHEN (ZTBPOTable.RD_DATE < GETDATE()-1) AND ((ZTBPOTable.PO_QTY-AA.TotalDelivered) <>0) THEN 'OVER'        ELSE 'OK' END AS OVERDUE, ZTBPOTable.REMARK, ZTBPOTable.PO_ID FROM (SELECT ZTBPOTable.EMPL_NO, ZTBPOTable.CUST_CD, ZTBPOTable.G_CODE, ZTBPOTable.PO_NO, isnull(SUM(ZTBDelivery.DELIVERY_QTY),0) AS TotalDelivered FROM ZTBPOTable  LEFT JOIN ZTBDelivery ON (ZTBDelivery.CTR_CD = ZTBPOTable.CTR_CD AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) GROUP BY ZTBPOTable.CTR_CD,ZTBPOTable.EMPL_NO,ZTBPOTable.G_CODE,ZTBPOTable.CUST_CD,ZTBPOTable.PO_NO) AS AA LEFT JOIN M010 ON (M010.EMPL_NO = AA.EMPL_NO) LEFT JOIN M100 ON (M100.G_CODE = AA.G_CODE) LEFT JOIN ZTBPOTable ON (AA.CUST_CD = ZTBPOTable.CUST_CD AND AA.G_CODE = ZTBPOTable.G_CODE AND AA.PO_NO = ZTBPOTable.PO_NO) JOIN M110 ON (M110.CUST_CD = AA.CUST_CD) ) AS XX WHERE XX.OVERDUE= 'OVER' GROUP BY XX.EMPL_NAME ORDER BY OVERDUE DESC";
            result = config.GetData(strQuery);
            return result;

        }


        public DataTable report_CustomerDeliveryByType(string startdate, string enddate)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT X1.CUST_NAME_KD, X1.DELIVERY_QTY, X1.TSP AS  TSP_QTY,X1.LABEL AS  LABEL_QTY,X1.UV AS  UV_QTY,X1.OLED AS  OLED_QTY,X1.TAPE AS  TAPE_QTY,X1.RIBBON AS  RIBBON_QTY,X1.SPT AS  SPT_QTY,X1.OTHERS AS  OTHERS_QTY, X2.DELIVERED_AMOUNT, X2.TSP AS  TSP_AMOUNT,X2.LABEL AS  LABEL_AMOUNT,X2.UV AS  UV_AMOUNT,X2.OLED AS  OLED_AMOUNT,X2.TAPE AS  TAPE_AMOUNT,X2.RIBBON AS  RIBBON_AMOUNT,X2.SPT AS  SPT_AMOUNT,X2.OTHERS AS  OTHERS_AMOUNT   FROM ( SELECT XX.CUST_NAME_KD, YY.DELIVERY_QTY, XX.TSP, XX.LABEL, XX.UV, XX.OLED, XX.TAPE, XX.RIBBON, XX. SPT, (YY.DELIVERY_QTY- XX.TSP- XX.LABEL- XX.UV- XX.OLED- XX.TAPE- XX.RIBBON- XX. SPT) AS OTHERS  FROM (SELECT PV.CUST_NAME_KD, ISNULL(PV.[TSP],0) As TSP,ISNULL(PV.[LABEL],0) As LABEL,ISNULL(PV.[UV],0) As UV,ISNULL(PV.[OLED],0) As OLED,ISNULL(PV.[TAPE],0) As TAPE,ISNULL(PV.[RIBBON],0) As RIBBON,ISNULL(PV.[SPT],0) As SPT FROM (SELECT AA.CUST_NAME_KD, AA.DELIVERY_QTY, AA.PROD_TYPE FROM (SELECT ZTBDelivery.G_CODE, M010.EMPL_NAME, M110.CUST_NAME_KD, ZTBDelivery.DELIVERY_DATE, M100.G_NAME, M100.PROD_MAIN_MATERIAL, ZTBDelivery.DELIVERY_QTY, ZTBPOTable.PROD_PRICE, ZTBDelivery.PO_NO, (ZTBPOTable.PROD_PRICE * ZTBDelivery.DELIVERY_QTY) As DELIVERED_AMOUNT, M100.PROD_TYPE, DATEPART( MONTH, ZTBDelivery.DELIVERY_DATE) AS DELMONTH, DATEPART( ISOWK,  ZTBDelivery.DELIVERY_DATE) AS DELWEEKNUM ,ZTBDelivery.NOCANCEL , ZTBDelivery.DELIVERY_ID  FROM ZTBDelivery JOIN ZTBPOTable ON (ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) JOIN M010 ON ZTBDelivery.EMPL_NO = M010.EMPL_NO JOIN M100 ON ZTBDelivery.G_CODE = M100.G_CODE JOIN M110 ON M110.CUST_CD = ZTBDelivery.CUST_CD) AS AA WHERE DELIVERY_DATE BETWEEN '" + startdate + "'  AND '" + enddate + "' ) AS BB PIVOT ( SUM(BB.DELIVERY_QTY) FOR BB.PROD_TYPE IN ([TSP],[LABEL],[UV],[OLED],[TAPE],[RIBBON],[SPT]) ) AS PV ) AS XX JOIN  (SELECT AA.CUST_NAME_KD, SUM(AA.DELIVERY_QTY) AS DELIVERY_QTY FROM (SELECT ZTBDelivery.G_CODE, M010.EMPL_NAME, M110.CUST_NAME_KD, ZTBDelivery.DELIVERY_DATE, M100.G_NAME, M100.PROD_MAIN_MATERIAL, ZTBDelivery.DELIVERY_QTY, ZTBPOTable.PROD_PRICE, ZTBDelivery.PO_NO, (ZTBPOTable.PROD_PRICE * ZTBDelivery.DELIVERY_QTY) As DELIVERED_AMOUNT, M100.PROD_TYPE, DATEPART( MONTH, ZTBDelivery.DELIVERY_DATE) AS DELMONTH, DATEPART( ISOWK,  ZTBDelivery.DELIVERY_DATE) AS DELWEEKNUM ,ZTBDelivery.NOCANCEL , ZTBDelivery.DELIVERY_ID  FROM ZTBDelivery JOIN ZTBPOTable ON (ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) JOIN M010 ON ZTBDelivery.EMPL_NO = M010.EMPL_NO JOIN M100 ON ZTBDelivery.G_CODE = M100.G_CODE JOIN M110 ON M110.CUST_CD = ZTBDelivery.CUST_CD) AS AA WHERE DELIVERY_DATE BETWEEN '" + startdate + "'  AND '" + enddate + "' GROUP BY AA.CUST_NAME_KD) AS YY ON XX.CUST_NAME_KD = YY.CUST_NAME_KD  ) X1 JOIN    /* CUSTOMER DELIVERY AMOUNT TABLE*/ (SELECT XX.CUST_NAME_KD, YY.DELIVERED_AMOUNT, XX.TSP, XX.LABEL, XX.UV, XX.OLED, XX.TAPE, XX.RIBBON, XX. SPT, (YY.DELIVERED_AMOUNT- XX.TSP- XX.LABEL- XX.UV- XX.OLED- XX.TAPE- XX.RIBBON- XX. SPT) AS OTHERS  FROM (SELECT PV.CUST_NAME_KD, ISNULL(PV.[TSP],0) As TSP,ISNULL(PV.[LABEL],0) As LABEL,ISNULL(PV.[UV],0) As UV,ISNULL(PV.[OLED],0) As OLED,ISNULL(PV.[TAPE],0) As TAPE,ISNULL(PV.[RIBBON],0) As RIBBON,ISNULL(PV.[SPT],0) As SPT FROM (SELECT AA.CUST_NAME_KD, AA.DELIVERED_AMOUNT, AA.PROD_TYPE FROM (SELECT ZTBDelivery.G_CODE, M010.EMPL_NAME, M110.CUST_NAME_KD, ZTBDelivery.DELIVERY_DATE, M100.G_NAME, M100.PROD_MAIN_MATERIAL, ZTBDelivery.DELIVERY_QTY, ZTBPOTable.PROD_PRICE, ZTBDelivery.PO_NO, (ZTBPOTable.PROD_PRICE * ZTBDelivery.DELIVERY_QTY) As DELIVERED_AMOUNT, M100.PROD_TYPE, DATEPART( MONTH, ZTBDelivery.DELIVERY_DATE) AS DELMONTH, DATEPART( ISOWK,  ZTBDelivery.DELIVERY_DATE) AS DELWEEKNUM ,ZTBDelivery.NOCANCEL , ZTBDelivery.DELIVERY_ID  FROM ZTBDelivery JOIN ZTBPOTable ON (ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) JOIN M010 ON ZTBDelivery.EMPL_NO = M010.EMPL_NO JOIN M100 ON ZTBDelivery.G_CODE = M100.G_CODE JOIN M110 ON M110.CUST_CD = ZTBDelivery.CUST_CD) AS AA WHERE DELIVERY_DATE BETWEEN '" + startdate + "'  AND '" + enddate + "' ) AS BB PIVOT ( SUM(BB.DELIVERED_AMOUNT) FOR BB.PROD_TYPE IN ([TSP],[LABEL],[UV],[OLED],[TAPE],[RIBBON],[SPT]) ) AS PV ) AS XX JOIN  (SELECT AA.CUST_NAME_KD, SUM(AA.DELIVERED_AMOUNT) AS DELIVERED_AMOUNT FROM (SELECT ZTBDelivery.G_CODE, M010.EMPL_NAME, M110.CUST_NAME_KD, ZTBDelivery.DELIVERY_DATE, M100.G_NAME, M100.PROD_MAIN_MATERIAL, ZTBDelivery.DELIVERY_QTY, ZTBPOTable.PROD_PRICE, ZTBDelivery.PO_NO, (ZTBPOTable.PROD_PRICE * ZTBDelivery.DELIVERY_QTY) As DELIVERED_AMOUNT, M100.PROD_TYPE, DATEPART( MONTH, ZTBDelivery.DELIVERY_DATE) AS DELMONTH, DATEPART( ISOWK,  ZTBDelivery.DELIVERY_DATE) AS DELWEEKNUM ,ZTBDelivery.NOCANCEL , ZTBDelivery.DELIVERY_ID  FROM ZTBDelivery JOIN ZTBPOTable ON (ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) JOIN M010 ON ZTBDelivery.EMPL_NO = M010.EMPL_NO JOIN M100 ON ZTBDelivery.G_CODE = M100.G_CODE JOIN M110 ON M110.CUST_CD = ZTBDelivery.CUST_CD) AS AA WHERE DELIVERY_DATE BETWEEN '" + startdate + "'  AND '" + enddate + "' GROUP BY AA.CUST_NAME_KD) AS YY ON XX.CUST_NAME_KD = YY.CUST_NAME_KD  ) AS X2 ON (X1.CUST_NAME_KD = X2.CUST_NAME_KD)";

            result = config.GetData(strQuery);
            return result;

        }


        public string report_SOP_uploadedPIC(string plandate)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT DISTINCT EMPL_NAME FROM  ZTBPLANTB JOIN M010 ON M010.EMPL_NO = ZTBPLANTB.EMPL_NO WHERE  PLAN_DATE = '" + plandate + "'";
            string picarray = "";            
            result = config.GetData(strQuery);
            if (result.Rows.Count > 0)
            {
                foreach(DataRow row in result.Rows)
                {
                    picarray = picarray + row[0].ToString() + "\n";
                }
            }
            else picarray = "NO";
            return picarray;

        }



        public DataTable report_SOP(string plandate)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "DECLARE @plandate date SET @plandate = '"+ plandate +"' SELECT @plandate AS PLAN_DATE ,II.G_CODE, M100.G_NAME_KD, M100.G_NAME, M100.PROD_TYPE, M100.PROD_MAIN_MATERIAL, II.PO_BALANCE,  II.D1,II.D2, II.D3, II.D4, II.D5, II.D6, II.D7, II.D8, CASE WHEN II.D9 < 0  THEN 0  ELSE D9 END AS D9 FROM (SELECT GG.G_CODE, GG.PO_BALANCE, isnull(HH.D1,0) AS D1 ,isnull(HH.D2,0) AS D2 , isnull(HH.D3,0) AS D3 , isnull(HH.D4,0) AS D4 , isnull(HH.D5,0) AS D5 , isnull(HH.D6,0) AS D6 , isnull(HH.D7,0) AS D7 , isnull(HH.D8,0) AS D8,  isnull((GG.PO_BALANCE- isnull( HH.D1,0)- isnull(HH.D2,0)- isnull( HH.D3,0)- isnull( HH.D4,0)- isnull( HH.D5,0)- isnull( HH.D6,0)- isnull( HH.D7,0)- isnull( HH.D8,0)),0) AS D9 FROM (SELECT  G_CODE, SUM(XX.PO_BALANCE) AS PO_BALANCE FROM (SELECT AA.PO_NO,  M100.PROD_TYPE, M100.PROD_MAIN_MATERIAL, ZTBPOTable.PO_DATE,ZTBPOTable.RD_DATE, M110.CUST_NAME_KD, M100.G_NAME, M010.EMPL_NAME, AA.G_CODE, ZTBPOTable.PO_QTY, ZTBPOTable.PROD_PRICE, AA.TotalDelivered as TOTAL_DELIVERED, (ZTBPOTable.PO_QTY-AA.TotalDelivered) As PO_BALANCE,(ZTBPOTable.PO_QTY*ZTBPOTable.PROD_PRICE) As PO_AMOUNT , (AA.TotalDelivered*ZTBPOTable.PROD_PRICE) As DELIVERED_AMOUNT, ((ZTBPOTable.PO_QTY-AA.TotalDelivered)*ZTBPOTable.PROD_PRICE) As BALANCE_AMOUNT,DATEPART( MONTH, PO_DATE) AS POMONTH, DATEPART( ISOWK, PO_DATE) AS POWEEKNUM, YEAR(PO_DATE) As PO_YEAR, CASE     WHEN (ZTBPOTable.RD_DATE < GETDATE()-1) AND ((ZTBPOTable.PO_QTY-AA.TotalDelivered) <>0) THEN 'OVER'        ELSE 'OK' END AS OVERDUE, ZTBPOTable.REMARK, ZTBPOTable.PO_ID FROM (SELECT ZTBPOTable.EMPL_NO, ZTBPOTable.CUST_CD, ZTBPOTable.G_CODE, ZTBPOTable.PO_NO, isnull(SUM(ZTBDelivery.DELIVERY_QTY),0) AS TotalDelivered FROM ZTBPOTable  LEFT JOIN ZTBDelivery ON (ZTBDelivery.CTR_CD = ZTBPOTable.CTR_CD AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) GROUP BY ZTBPOTable.CTR_CD,ZTBPOTable.EMPL_NO,ZTBPOTable.G_CODE,ZTBPOTable.CUST_CD,ZTBPOTable.PO_NO) AS AA LEFT JOIN M010 ON (M010.EMPL_NO = AA.EMPL_NO) LEFT JOIN M100 ON (M100.G_CODE = AA.G_CODE) LEFT JOIN ZTBPOTable ON (AA.CUST_CD = ZTBPOTable.CUST_CD AND AA.G_CODE = ZTBPOTable.G_CODE AND AA.PO_NO = ZTBPOTable.PO_NO) JOIN M110 ON (M110.CUST_CD = AA.CUST_CD) ) AS XX GROUP BY G_CODE ) AS GG LEFT OUTER JOIN ( SELECT DISTINCT G_CODE, SUM(D1)AS D1,SUM(D2)AS D2,SUM(D3)AS D3, SUM(D4)AS D4, SUM(D5)AS D5, SUM(D6)AS D6,SUM(D7)AS D7,SUM(D8) AS D8 FROM ZTBPLANTB WHERE PLAN_DATE= @plandate GROUP BY G_CODE ) AS HH ON GG.G_CODE = HH.G_CODE) AS II JOIN M100 ON M100.G_CODE = II.G_CODE";
            result = config.GetData(strQuery);
            return result;

        }


        public DataTable report_CustomerPOBalanceByType()
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT XX.CUST_NAME_KD, YY.TOTAL_PO_BALANCE, XX.TSP, XX.LABEL, XX.UV, XX.OLED, XX.TAPE, XX.RIBBON, XX.SPT, (YY.TOTAL_PO_BALANCE- XX.TSP- XX.LABEL- XX.UV- XX.OLED- XX.TAPE- XX.RIBBON- XX.SPT) AS OTHERS FROM  (SELECT  PV.CUST_NAME_KD, (isnull(PV.[TSP],0)+isnull(PV.[LABEL],0)+isnull(PV.[UV],0)+isnull(PV.[TAPE],0) + isnull(PV.[SPT],0)+ isnull(PV.[OLED],0)  + isnull(PV.[RIBBON],0)) As TOTAL_PO_BALANCE, isnull(PV.[TSP],0) As TSP, isnull(PV.[LABEL],0) As LABEL,isnull(PV.[UV],0) As UV, isnull(PV.[OLED],0) As OLED,isnull(PV.[TAPE],0) As TAPE, isnull(PV.[SPT],0) As SPT, isnull(PV.[RIBBON],0) As RIBBON FROM ( SELECT P.PO_BALANCE, P.PROD_TYPE, P.CUST_NAME_KD FROM (   SELECT AA.PO_NO,  M100.PROD_TYPE, M100.PROD_MAIN_MATERIAL, ZTBPOTable.PO_DATE,ZTBPOTable.RD_DATE, M110.CUST_NAME_KD, M100.G_NAME, M010.EMPL_NAME, AA.G_CODE, ZTBPOTable.PO_QTY, ZTBPOTable.PROD_PRICE, AA.TotalDelivered as TOTAL_DELIVERED, (ZTBPOTable.PO_QTY-AA.TotalDelivered) As PO_BALANCE,(ZTBPOTable.PO_QTY*ZTBPOTable.PROD_PRICE) As PO_AMOUNT , (AA.TotalDelivered*ZTBPOTable.PROD_PRICE) As DELIVERED_AMOUNT, ((ZTBPOTable.PO_QTY-AA.TotalDelivered)*ZTBPOTable.PROD_PRICE) As BALANCE_AMOUNT,DATEPART( MONTH, PO_DATE) AS POMONTH, DATEPART( ISOWK, PO_DATE) AS POWEEKNUM, YEAR(PO_DATE) As PO_YEAR, CASE     WHEN (ZTBPOTable.RD_DATE < GETDATE()-1) AND ((ZTBPOTable.PO_QTY-AA.TotalDelivered) <>0) THEN 'OVER'        ELSE 'OK' END AS OVERDUE, ZTBPOTable.REMARK, ZTBPOTable.PO_ID FROM (SELECT ZTBPOTable.EMPL_NO, ZTBPOTable.CUST_CD, ZTBPOTable.G_CODE, ZTBPOTable.PO_NO, isnull(SUM(ZTBDelivery.DELIVERY_QTY),0) AS TotalDelivered FROM ZTBPOTable  LEFT JOIN ZTBDelivery ON (ZTBDelivery.CTR_CD = ZTBPOTable.CTR_CD AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) GROUP BY ZTBPOTable.CTR_CD,ZTBPOTable.EMPL_NO,ZTBPOTable.G_CODE,ZTBPOTable.CUST_CD,ZTBPOTable.PO_NO) AS AA LEFT JOIN M010 ON (M010.EMPL_NO = AA.EMPL_NO) LEFT JOIN M100 ON (M100.G_CODE = AA.G_CODE) LEFT JOIN ZTBPOTable ON (AA.CUST_CD = ZTBPOTable.CUST_CD AND AA.G_CODE = ZTBPOTable.G_CODE AND AA.PO_NO = ZTBPOTable.PO_NO) JOIN M110 ON (M110.CUST_CD = AA.CUST_CD)  ) AS P ) AS j PIVOT (SUM(j.PO_BALANCE) FOR j.PROD_TYPE IN ([TSP],[LABEL],[UV],[OLED],[TAPE],[SPT],[RIBBON])) AS PV ) AS XX JOIN /*customer sum po balance*/  (SELECT AA.CUST_NAME_KD, SUM(AA.PO_BALANCE) AS TOTAL_PO_BALANCE FROM (SELECT AA.PO_NO,  M100.PROD_TYPE, M100.PROD_MAIN_MATERIAL, ZTBPOTable.PO_DATE,ZTBPOTable.RD_DATE, M110.CUST_NAME_KD, M100.G_NAME, M010.EMPL_NAME, AA.G_CODE, ZTBPOTable.PO_QTY, ZTBPOTable.PROD_PRICE, AA.TotalDelivered as TOTAL_DELIVERED, (ZTBPOTable.PO_QTY-AA.TotalDelivered) As PO_BALANCE,(ZTBPOTable.PO_QTY*ZTBPOTable.PROD_PRICE) As PO_AMOUNT , (AA.TotalDelivered*ZTBPOTable.PROD_PRICE) As DELIVERED_AMOUNT, ((ZTBPOTable.PO_QTY-AA.TotalDelivered)*ZTBPOTable.PROD_PRICE) As BALANCE_AMOUNT,DATEPART( MONTH, PO_DATE) AS POMONTH, DATEPART( ISOWK, PO_DATE) AS POWEEKNUM, YEAR(PO_DATE) As PO_YEAR, CASE     WHEN (ZTBPOTable.RD_DATE < GETDATE()-1) AND ((ZTBPOTable.PO_QTY-AA.TotalDelivered) <>0) THEN 'OVER'        ELSE 'OK' END AS OVERDUE, ZTBPOTable.REMARK, ZTBPOTable.PO_ID FROM (SELECT ZTBPOTable.EMPL_NO, ZTBPOTable.CUST_CD, ZTBPOTable.G_CODE, ZTBPOTable.PO_NO, isnull(SUM(ZTBDelivery.DELIVERY_QTY),0) AS TotalDelivered FROM ZTBPOTable  LEFT JOIN ZTBDelivery ON (ZTBDelivery.CTR_CD = ZTBPOTable.CTR_CD AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) GROUP BY ZTBPOTable.CTR_CD,ZTBPOTable.EMPL_NO,ZTBPOTable.G_CODE,ZTBPOTable.CUST_CD,ZTBPOTable.PO_NO) AS AA LEFT JOIN M010 ON (M010.EMPL_NO = AA.EMPL_NO) LEFT JOIN M100 ON (M100.G_CODE = AA.G_CODE) LEFT JOIN ZTBPOTable ON (AA.CUST_CD = ZTBPOTable.CUST_CD AND AA.G_CODE = ZTBPOTable.G_CODE AND AA.PO_NO = ZTBPOTable.PO_NO) JOIN M110 ON (M110.CUST_CD = AA.CUST_CD)) AS AA GROUP BY AA.CUST_NAME_KD ) AS YY ON XX.CUST_NAME_KD = YY.CUST_NAME_KD ORDER BY TOTAL_PO_BALANCE DESC";
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable report_CustomerFcstByWeek(int weekNo, int fcstyear)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "USE CMS_VINA DECLARE @tuan AS int DECLARE @nam AS int  SET @tuan = " + weekNo + " SET @nam = " + fcstyear + " SELECT   @tuan AS WEEKNUM , M110.CUST_NAME_KD,SUM(ZTBFCSTTB.W1)AS W1 , SUM(ZTBFCSTTB.W2)AS W2 , SUM(ZTBFCSTTB.W3)AS W3 , SUM(ZTBFCSTTB.W4)AS W4 , SUM(ZTBFCSTTB.W5)AS W5 , SUM(ZTBFCSTTB.W6)AS W6 , SUM(ZTBFCSTTB.W7)AS W7 , SUM(ZTBFCSTTB.W8)AS W8 , SUM(ZTBFCSTTB.W9)AS W9 , SUM(ZTBFCSTTB.W10)AS W10 , SUM(ZTBFCSTTB.W11)AS W11   FROM ZTBFCSTTB JOIN M100 ON (M100.G_CODE = ZTBFCSTTB.G_CODE) JOIN M110 ON(M110.CUST_CD = ZTBFCSTTB.CUST_CD) JOIN M010 ON (M010.EMPL_NO = ZTBFCSTTB.EMPL_NO) WHERE (CUST_NAME_KD = 'SEV' OR CUST_NAME_KD = 'SEVT' OR CUST_NAME_KD = 'SAMSUNG-ASIA') AND (FCSTWEEKNO = @tuan) AND (FCSTYEAR = @nam) GROUP BY M110.CUST_NAME_KD ORDER BY M110.CUST_NAME_KD DESC";
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable report_POBalanceByTypeSS(string balanceDate)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            //string strQuery = "USE CMS_VINA DECLARE @ngaythang AS DATE SET @ngaythang = '" + balanceDate + "'  SELECT YEAR(@ngaythang) AS BALANCE_YEAR, DATEPART( ISOWK, @ngaythang) AS BALANCE_WEEKNUM ,@ngaythang AS BALANCE_DATE, SUM(TOTAL_PO_BALANCE) AS TOTAL_PO_BALANCE, SUM(TSP) AS TSP, SUM(LABEL) AS LABEL, SUM(UV) AS UV,SUM(OLED) AS OLED,SUM(TAPE) AS TAPE, SUM(RIBBON) AS RIBBON,SUM(ALUMIUM) AS ALUMIUM, (SUM(TOTAL_PO_BALANCE) -  SUM(TSP) -SUM(LABEL) -SUM(UV) - SUM(OLED) -SUM(TAPE) -SUM(RIBBON) -SUM(ALUMIUM)) AS OTHERS FROM (  	SELECT AA.POWEEKNUM, BB.TOTAL_PO_BALANCE, AA.TSP, AA.LABEL, AA.UV, AA.OLED, AA.TAPE, AA.RIBBON, AA.ALUMIUM, (BB.TOTAL_PO_BALANCE-AA.TSP- AA.LABEL- AA.UV- AA.OLED- AA.TAPE- AA.RIBBON- AA.ALUMIUM) AS OTHERS FROM 	( 		SELECT  POWEEKNUM,   isnull(PV.[TSP],0) As TSP, isnull(PV.[LABEL],0) As LABEL, isnull(PV.[UV],0) As UV,isnull(PV.[OLED],0) As OLED,isnull(PV.[TAPE],0) As TAPE,isnull(PV.[RIBBON],0) As RIBBON,isnull(PV.[ALUMIUM],0) As ALUMIUM 		FROM 		( 		SELECT P.PO_BALANCE, P.PROD_TYPE,  P.POWEEKNUM FROM 			( 	 				SELECT AA.PO_NO,  M100.PROD_TYPE, M100.PROD_MAIN_MATERIAL, ZTBPOTable.PO_DATE,ZTBPOTable.RD_DATE, M110.CUST_NAME_KD, M100.G_NAME, M010.EMPL_NAME, AA.G_CODE, ZTBPOTable.PO_QTY, ZTBPOTable.PROD_PRICE, AA.TotalDelivered as TOTAL_DELIVERED, (ZTBPOTable.PO_QTY-AA.TotalDelivered) As PO_BALANCE,(ZTBPOTable.PO_QTY*ZTBPOTable.PROD_PRICE) As PO_AMOUNT , (AA.TotalDelivered*ZTBPOTable.PROD_PRICE) As DELIVERED_AMOUNT, ((ZTBPOTable.PO_QTY-AA.TotalDelivered)*ZTBPOTable.PROD_PRICE) As BALANCE_AMOUNT,DATEPART( MONTH, PO_DATE) AS POMONTH, DATEPART( ISOWK, PO_DATE) AS POWEEKNUM, 				CASE 					WHEN (ZTBPOTable.RD_DATE < GETDATE()-1) AND ((ZTBPOTable.PO_QTY-AA.TotalDelivered) <>0) THEN 'OVER'    					ELSE 'OK' 				END AS OVERDUE, ZTBPOTable.REMARK, ZTBPOTable.PO_ID 				FROM 				(SELECT ZTBPOTable.EMPL_NO, ZTBPOTable.CUST_CD, ZTBPOTable.G_CODE, ZTBPOTable.PO_NO, isnull(SUM(ZTBDelivery.DELIVERY_QTY),0) AS TotalDelivered 				FROM 				fn_PODATE(@ngaythang)  AS ZTBPOTable 				LEFT JOIN 				fn_DELIVERYTODATE(@ngaythang) AS  ZTBDelivery 				ON (ZTBDelivery.CTR_CD = ZTBPOTable.CTR_CD AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) 				GROUP BY ZTBPOTable.CTR_CD,ZTBPOTable.EMPL_NO,ZTBPOTable.G_CODE,ZTBPOTable.CUST_CD,ZTBPOTable.PO_NO) AS AA 				LEFT JOIN M010 				ON (M010.EMPL_NO = AA.EMPL_NO) 				LEFT JOIN M100 				ON (M100.G_CODE = AA.G_CODE) 				LEFT JOIN ZTBPOTable 				ON (AA.CUST_CD = ZTBPOTable.CUST_CD AND AA.G_CODE = ZTBPOTable.G_CODE AND AA.PO_NO = ZTBPOTable.PO_NO) 				JOIN M110 				ON (M110.CUST_CD = AA.CUST_CD) 				WHERE CUST_NAME_KD='SEVT' OR  CUST_NAME_KD='SEV' OR  CUST_NAME_KD='SAMSUNG-ASIA'	 	  			) AS P 		) AS j 		PIVOT (SUM(j.PO_BALANCE) FOR j.PROD_TYPE IN ([TSP],[LABEL],[UV],[OLED],[TAPE],[RIBBON],[ALUMIUM])) AS PV 	) AS AA 	JOIN 	(   	SELECT A.POWEEKNUM, SUM(A.PO_BALANCE) As TOTAL_PO_BALANCE 	FROM 	(   		SELECT AA.PO_NO,  M100.PROD_TYPE, M100.PROD_MAIN_MATERIAL, ZTBPOTable.PO_DATE,ZTBPOTable.RD_DATE, M110.CUST_NAME_KD, M100.G_NAME, M010.EMPL_NAME, AA.G_CODE, ZTBPOTable.PO_QTY, ZTBPOTable.PROD_PRICE, AA.TotalDelivered as TOTAL_DELIVERED, (ZTBPOTable.PO_QTY-AA.TotalDelivered) As PO_BALANCE,(ZTBPOTable.PO_QTY*ZTBPOTable.PROD_PRICE) As PO_AMOUNT , (AA.TotalDelivered*ZTBPOTable.PROD_PRICE) As DELIVERED_AMOUNT, ((ZTBPOTable.PO_QTY-AA.TotalDelivered)*ZTBPOTable.PROD_PRICE) As BALANCE_AMOUNT,DATEPART( MONTH, PO_DATE) AS POMONTH, DATEPART( ISOWK, PO_DATE) AS POWEEKNUM, 				CASE 					WHEN (ZTBPOTable.RD_DATE < GETDATE()-1) AND ((ZTBPOTable.PO_QTY-AA.TotalDelivered) <>0) THEN 'OVER'    					ELSE 'OK' 				END AS OVERDUE, ZTBPOTable.REMARK, ZTBPOTable.PO_ID 				FROM 				(SELECT ZTBPOTable.EMPL_NO, ZTBPOTable.CUST_CD, ZTBPOTable.G_CODE, ZTBPOTable.PO_NO, isnull(SUM(ZTBDelivery.DELIVERY_QTY),0) AS TotalDelivered 				FROM 				fn_PODATE(@ngaythang)  AS ZTBPOTable 				LEFT JOIN 				fn_DELIVERYTODATE(@ngaythang) AS  ZTBDelivery 				ON (ZTBDelivery.CTR_CD = ZTBPOTable.CTR_CD AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) 				GROUP BY ZTBPOTable.CTR_CD,ZTBPOTable.EMPL_NO,ZTBPOTable.G_CODE,ZTBPOTable.CUST_CD,ZTBPOTable.PO_NO) AS AA 				LEFT JOIN M010 				ON (M010.EMPL_NO = AA.EMPL_NO) 				LEFT JOIN M100 				ON (M100.G_CODE = AA.G_CODE) 				LEFT JOIN ZTBPOTable 				ON (AA.CUST_CD = ZTBPOTable.CUST_CD AND AA.G_CODE = ZTBPOTable.G_CODE AND AA.PO_NO = ZTBPOTable.PO_NO) 				JOIN M110 				ON (M110.CUST_CD = AA.CUST_CD) 				WHERE CUST_NAME_KD='SEVT' OR  CUST_NAME_KD='SEV' OR  CUST_NAME_KD='SAMSUNG-ASIA'	  	) AS A 	GROUP BY POWEEKNUM 	) AS BB 	ON AA.POWEEKNUM = BB.POWEEKNUM ) AS XX";
            string strQuery = "USE CMS_VINA DECLARE @ngaythang AS DATE SET @ngaythang = '" + balanceDate + "'  SELECT YEAR(@ngaythang) AS BALANCE_YEAR, DATEPART( ISOWK, @ngaythang) AS BALANCE_WEEKNUM ,@ngaythang AS BALANCE_DATE, SUM(TOTAL_PO_BALANCE) AS TOTAL_PO_BALANCE, SUM(TSP) AS TSP, SUM(LABEL) AS LABEL, SUM(UV) AS UV,SUM(OLED) AS OLED,SUM(TAPE) AS TAPE, SUM(RIBBON) AS RIBBON,SUM(SPT) AS SPT, (SUM(TOTAL_PO_BALANCE) -  SUM(TSP) -SUM(LABEL) -SUM(UV) - SUM(OLED) -SUM(TAPE) -SUM(RIBBON) -SUM(SPT)) AS OTHERS FROM (  	SELECT AA.POWEEKNUM, BB.TOTAL_PO_BALANCE, AA.TSP, AA.LABEL, AA.UV, AA.OLED, AA.TAPE, AA.RIBBON, AA.SPT, (BB.TOTAL_PO_BALANCE-AA.TSP- AA.LABEL- AA.UV- AA.OLED- AA.TAPE- AA.RIBBON- AA.SPT) AS OTHERS FROM 	( 		SELECT  POWEEKNUM,   isnull(PV.[TSP],0) As TSP, isnull(PV.[LABEL],0) As LABEL, isnull(PV.[UV],0) As UV,isnull(PV.[OLED],0) As OLED,isnull(PV.[TAPE],0) As TAPE,isnull(PV.[RIBBON],0) As RIBBON,isnull(PV.[SPT],0) As SPT 		FROM 		( 		SELECT P.PO_BALANCE, P.PROD_TYPE,  P.POWEEKNUM FROM 			( 	 				SELECT AA.PO_NO,  M100.PROD_TYPE, M100.PROD_MAIN_MATERIAL, ZTBPOTable.PO_DATE,ZTBPOTable.RD_DATE, M110.CUST_NAME_KD, M100.G_NAME, M010.EMPL_NAME, AA.G_CODE, ZTBPOTable.PO_QTY, ZTBPOTable.PROD_PRICE, AA.TotalDelivered as TOTAL_DELIVERED, (ZTBPOTable.PO_QTY-AA.TotalDelivered) As PO_BALANCE,(ZTBPOTable.PO_QTY*ZTBPOTable.PROD_PRICE) As PO_AMOUNT , (AA.TotalDelivered*ZTBPOTable.PROD_PRICE) As DELIVERED_AMOUNT, ((ZTBPOTable.PO_QTY-AA.TotalDelivered)*ZTBPOTable.PROD_PRICE) As BALANCE_AMOUNT,DATEPART( MONTH, PO_DATE) AS POMONTH, DATEPART( ISOWK, PO_DATE) AS POWEEKNUM, 				CASE 					WHEN (ZTBPOTable.RD_DATE < GETDATE()-1) AND ((ZTBPOTable.PO_QTY-AA.TotalDelivered) <>0) THEN 'OVER'    					ELSE 'OK' 				END AS OVERDUE, ZTBPOTable.REMARK, ZTBPOTable.PO_ID 				FROM 				(SELECT ZTBPOTable.EMPL_NO, ZTBPOTable.CUST_CD, ZTBPOTable.G_CODE, ZTBPOTable.PO_NO, isnull(SUM(ZTBDelivery.DELIVERY_QTY),0) AS TotalDelivered 				FROM 				fn_PODATE(@ngaythang)  AS ZTBPOTable 				LEFT JOIN 				fn_DELIVERYTODATE(@ngaythang) AS  ZTBDelivery 				ON (ZTBDelivery.CTR_CD = ZTBPOTable.CTR_CD AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) 				GROUP BY ZTBPOTable.CTR_CD,ZTBPOTable.EMPL_NO,ZTBPOTable.G_CODE,ZTBPOTable.CUST_CD,ZTBPOTable.PO_NO) AS AA 				LEFT JOIN M010 				ON (M010.EMPL_NO = AA.EMPL_NO) 				LEFT JOIN M100 				ON (M100.G_CODE = AA.G_CODE) 				LEFT JOIN ZTBPOTable 				ON (AA.CUST_CD = ZTBPOTable.CUST_CD AND AA.G_CODE = ZTBPOTable.G_CODE AND AA.PO_NO = ZTBPOTable.PO_NO) 				JOIN M110 				ON (M110.CUST_CD = AA.CUST_CD) 				WHERE CUST_NAME_KD='SEVT' OR  CUST_NAME_KD='SEV' OR  CUST_NAME_KD='SAMSUNG-ASIA'	 	  			) AS P 		) AS j 		PIVOT (SUM(j.PO_BALANCE) FOR j.PROD_TYPE IN ([TSP],[LABEL],[UV],[OLED],[TAPE],[RIBBON],[SPT])) AS PV 	) AS AA 	JOIN 	(   	SELECT A.POWEEKNUM, SUM(A.PO_BALANCE) As TOTAL_PO_BALANCE 	FROM 	(   		SELECT AA.PO_NO,  M100.PROD_TYPE, M100.PROD_MAIN_MATERIAL, ZTBPOTable.PO_DATE,ZTBPOTable.RD_DATE, M110.CUST_NAME_KD, M100.G_NAME, M010.EMPL_NAME, AA.G_CODE, ZTBPOTable.PO_QTY, ZTBPOTable.PROD_PRICE, AA.TotalDelivered as TOTAL_DELIVERED, (ZTBPOTable.PO_QTY-AA.TotalDelivered) As PO_BALANCE,(ZTBPOTable.PO_QTY*ZTBPOTable.PROD_PRICE) As PO_AMOUNT , (AA.TotalDelivered*ZTBPOTable.PROD_PRICE) As DELIVERED_AMOUNT, ((ZTBPOTable.PO_QTY-AA.TotalDelivered)*ZTBPOTable.PROD_PRICE) As BALANCE_AMOUNT,DATEPART( MONTH, PO_DATE) AS POMONTH, DATEPART( ISOWK, PO_DATE) AS POWEEKNUM, 				CASE 					WHEN (ZTBPOTable.RD_DATE < GETDATE()-1) AND ((ZTBPOTable.PO_QTY-AA.TotalDelivered) <>0) THEN 'OVER'    					ELSE 'OK' 				END AS OVERDUE, ZTBPOTable.REMARK, ZTBPOTable.PO_ID 				FROM 				(SELECT ZTBPOTable.EMPL_NO, ZTBPOTable.CUST_CD, ZTBPOTable.G_CODE, ZTBPOTable.PO_NO, isnull(SUM(ZTBDelivery.DELIVERY_QTY),0) AS TotalDelivered 				FROM 				fn_PODATE(@ngaythang)  AS ZTBPOTable 				LEFT JOIN 				fn_DELIVERYTODATE(@ngaythang) AS  ZTBDelivery 				ON (ZTBDelivery.CTR_CD = ZTBPOTable.CTR_CD AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) 				GROUP BY ZTBPOTable.CTR_CD,ZTBPOTable.EMPL_NO,ZTBPOTable.G_CODE,ZTBPOTable.CUST_CD,ZTBPOTable.PO_NO) AS AA 				LEFT JOIN M010 				ON (M010.EMPL_NO = AA.EMPL_NO) 				LEFT JOIN M100 				ON (M100.G_CODE = AA.G_CODE) 				LEFT JOIN ZTBPOTable 				ON (AA.CUST_CD = ZTBPOTable.CUST_CD AND AA.G_CODE = ZTBPOTable.G_CODE AND AA.PO_NO = ZTBPOTable.PO_NO) 				JOIN M110 				ON (M110.CUST_CD = AA.CUST_CD) 				WHERE CUST_NAME_KD='SEVT' OR  CUST_NAME_KD='SEV' OR  CUST_NAME_KD='SAMSUNG-ASIA'	  	) AS A 	GROUP BY POWEEKNUM 	) AS BB 	ON AA.POWEEKNUM = BB.POWEEKNUM ) AS XX";
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable report_POBalanceByType(string balanceDate)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            //string strQuery = "USE CMS_VINA DECLARE @ngaythang AS DATE SET @ngaythang = '" + balanceDate + "'  SELECT YEAR(@ngaythang) AS BALANCE_YEAR, DATEPART( ISOWK, @ngaythang) AS BALANCE_WEEKNUM ,@ngaythang AS BALANCE_DATE, SUM(TOTAL_PO_BALANCE) AS TOTAL_PO_BALANCE, SUM(TSP) AS TSP, SUM(LABEL) AS LABEL, SUM(UV) AS UV,SUM(OLED) AS OLED,SUM(TAPE) AS TAPE, SUM(RIBBON) AS RIBBON,SUM(ALUMIUM) AS ALUMIUM, (SUM(TOTAL_PO_BALANCE) -  SUM(TSP) -SUM(LABEL) -SUM(UV) - SUM(OLED) -SUM(TAPE) -SUM(RIBBON) -SUM(ALUMIUM)) AS OTHERS FROM (  	SELECT AA.POWEEKNUM, BB.TOTAL_PO_BALANCE, AA.TSP, AA.LABEL, AA.UV, AA.OLED, AA.TAPE, AA.RIBBON, AA.ALUMIUM, (BB.TOTAL_PO_BALANCE-AA.TSP- AA.LABEL- AA.UV- AA.OLED- AA.TAPE- AA.RIBBON- AA.ALUMIUM) AS OTHERS FROM 	( 		SELECT  POWEEKNUM,   isnull(PV.[TSP],0) As TSP, isnull(PV.[LABEL],0) As LABEL, isnull(PV.[UV],0) As UV,isnull(PV.[OLED],0) As OLED,isnull(PV.[TAPE],0) As TAPE,isnull(PV.[RIBBON],0) As RIBBON,isnull(PV.[ALUMIUM],0) As ALUMIUM 		FROM 		( 		SELECT P.PO_BALANCE, P.PROD_TYPE,  P.POWEEKNUM FROM 			( 	 				SELECT AA.PO_NO,  M100.PROD_TYPE, M100.PROD_MAIN_MATERIAL, ZTBPOTable.PO_DATE,ZTBPOTable.RD_DATE, M110.CUST_NAME_KD, M100.G_NAME, M010.EMPL_NAME, AA.G_CODE, ZTBPOTable.PO_QTY, ZTBPOTable.PROD_PRICE, AA.TotalDelivered as TOTAL_DELIVERED, (ZTBPOTable.PO_QTY-AA.TotalDelivered) As PO_BALANCE,(ZTBPOTable.PO_QTY*ZTBPOTable.PROD_PRICE) As PO_AMOUNT , (AA.TotalDelivered*ZTBPOTable.PROD_PRICE) As DELIVERED_AMOUNT, ((ZTBPOTable.PO_QTY-AA.TotalDelivered)*ZTBPOTable.PROD_PRICE) As BALANCE_AMOUNT,DATEPART( MONTH, PO_DATE) AS POMONTH, DATEPART( ISOWK, PO_DATE) AS POWEEKNUM, 				CASE 					WHEN (ZTBPOTable.RD_DATE < GETDATE()-1) AND ((ZTBPOTable.PO_QTY-AA.TotalDelivered) <>0) THEN 'OVER'    					ELSE 'OK' 				END AS OVERDUE, ZTBPOTable.REMARK, ZTBPOTable.PO_ID 				FROM 				(SELECT ZTBPOTable.EMPL_NO, ZTBPOTable.CUST_CD, ZTBPOTable.G_CODE, ZTBPOTable.PO_NO, isnull(SUM(ZTBDelivery.DELIVERY_QTY),0) AS TotalDelivered 				FROM 				fn_PODATE(@ngaythang)  AS ZTBPOTable 				LEFT JOIN 				fn_DELIVERYTODATE(@ngaythang) AS  ZTBDelivery 				ON (ZTBDelivery.CTR_CD = ZTBPOTable.CTR_CD AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) 				GROUP BY ZTBPOTable.CTR_CD,ZTBPOTable.EMPL_NO,ZTBPOTable.G_CODE,ZTBPOTable.CUST_CD,ZTBPOTable.PO_NO) AS AA 				LEFT JOIN M010 				ON (M010.EMPL_NO = AA.EMPL_NO) 				LEFT JOIN M100 				ON (M100.G_CODE = AA.G_CODE) 				LEFT JOIN ZTBPOTable 				ON (AA.CUST_CD = ZTBPOTable.CUST_CD AND AA.G_CODE = ZTBPOTable.G_CODE AND AA.PO_NO = ZTBPOTable.PO_NO) 				JOIN M110 				ON (M110.CUST_CD = AA.CUST_CD) 				WHERE CUST_NAME_KD='SEVT' OR  CUST_NAME_KD='SEV' OR  CUST_NAME_KD='SAMSUNG-ASIA'	 	  			) AS P 		) AS j 		PIVOT (SUM(j.PO_BALANCE) FOR j.PROD_TYPE IN ([TSP],[LABEL],[UV],[OLED],[TAPE],[RIBBON],[ALUMIUM])) AS PV 	) AS AA 	JOIN 	(   	SELECT A.POWEEKNUM, SUM(A.PO_BALANCE) As TOTAL_PO_BALANCE 	FROM 	(   		SELECT AA.PO_NO,  M100.PROD_TYPE, M100.PROD_MAIN_MATERIAL, ZTBPOTable.PO_DATE,ZTBPOTable.RD_DATE, M110.CUST_NAME_KD, M100.G_NAME, M010.EMPL_NAME, AA.G_CODE, ZTBPOTable.PO_QTY, ZTBPOTable.PROD_PRICE, AA.TotalDelivered as TOTAL_DELIVERED, (ZTBPOTable.PO_QTY-AA.TotalDelivered) As PO_BALANCE,(ZTBPOTable.PO_QTY*ZTBPOTable.PROD_PRICE) As PO_AMOUNT , (AA.TotalDelivered*ZTBPOTable.PROD_PRICE) As DELIVERED_AMOUNT, ((ZTBPOTable.PO_QTY-AA.TotalDelivered)*ZTBPOTable.PROD_PRICE) As BALANCE_AMOUNT,DATEPART( MONTH, PO_DATE) AS POMONTH, DATEPART( ISOWK, PO_DATE) AS POWEEKNUM, 				CASE 					WHEN (ZTBPOTable.RD_DATE < GETDATE()-1) AND ((ZTBPOTable.PO_QTY-AA.TotalDelivered) <>0) THEN 'OVER'    					ELSE 'OK' 				END AS OVERDUE, ZTBPOTable.REMARK, ZTBPOTable.PO_ID 				FROM 				(SELECT ZTBPOTable.EMPL_NO, ZTBPOTable.CUST_CD, ZTBPOTable.G_CODE, ZTBPOTable.PO_NO, isnull(SUM(ZTBDelivery.DELIVERY_QTY),0) AS TotalDelivered 				FROM 				fn_PODATE(@ngaythang)  AS ZTBPOTable 				LEFT JOIN 				fn_DELIVERYTODATE(@ngaythang) AS  ZTBDelivery 				ON (ZTBDelivery.CTR_CD = ZTBPOTable.CTR_CD AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) 				GROUP BY ZTBPOTable.CTR_CD,ZTBPOTable.EMPL_NO,ZTBPOTable.G_CODE,ZTBPOTable.CUST_CD,ZTBPOTable.PO_NO) AS AA 				LEFT JOIN M010 				ON (M010.EMPL_NO = AA.EMPL_NO) 				LEFT JOIN M100 				ON (M100.G_CODE = AA.G_CODE) 				LEFT JOIN ZTBPOTable 				ON (AA.CUST_CD = ZTBPOTable.CUST_CD AND AA.G_CODE = ZTBPOTable.G_CODE AND AA.PO_NO = ZTBPOTable.PO_NO) 				JOIN M110 				ON (M110.CUST_CD = AA.CUST_CD) 				WHERE CUST_NAME_KD='SEVT' OR  CUST_NAME_KD='SEV' OR  CUST_NAME_KD='SAMSUNG-ASIA'	  	) AS A 	GROUP BY POWEEKNUM 	) AS BB 	ON AA.POWEEKNUM = BB.POWEEKNUM ) AS XX";
            string strQuery = "USE CMS_VINA DECLARE @ngaythang AS DATE SET @ngaythang = '" + balanceDate + "'  SELECT YEAR(@ngaythang) AS BALANCE_YEAR, DATEPART( ISOWK, @ngaythang) AS BALANCE_WEEKNUM ,@ngaythang AS BALANCE_DATE, SUM(TOTAL_PO_BALANCE) AS TOTAL_PO_BALANCE, SUM(TSP) AS TSP, SUM(LABEL) AS LABEL, SUM(UV) AS UV,SUM(OLED) AS OLED,SUM(TAPE) AS TAPE, SUM(RIBBON) AS RIBBON,SUM(SPT) AS SPT, (SUM(TOTAL_PO_BALANCE) -  SUM(TSP) -SUM(LABEL) -SUM(UV) - SUM(OLED) -SUM(TAPE) -SUM(RIBBON) -SUM(SPT)) AS OTHERS FROM (  	SELECT AA.POWEEKNUM, BB.TOTAL_PO_BALANCE, AA.TSP, AA.LABEL, AA.UV, AA.OLED, AA.TAPE, AA.RIBBON, AA.SPT, (BB.TOTAL_PO_BALANCE-AA.TSP- AA.LABEL- AA.UV- AA.OLED- AA.TAPE- AA.RIBBON- AA.SPT) AS OTHERS FROM 	( 		SELECT  POWEEKNUM,   isnull(PV.[TSP],0) As TSP, isnull(PV.[LABEL],0) As LABEL, isnull(PV.[UV],0) As UV,isnull(PV.[OLED],0) As OLED,isnull(PV.[TAPE],0) As TAPE,isnull(PV.[RIBBON],0) As RIBBON,isnull(PV.[SPT],0) As SPT 		FROM 		( 		SELECT P.PO_BALANCE, P.PROD_TYPE,  P.POWEEKNUM FROM 			( 	 				SELECT AA.PO_NO,  M100.PROD_TYPE, M100.PROD_MAIN_MATERIAL, ZTBPOTable.PO_DATE,ZTBPOTable.RD_DATE, M110.CUST_NAME_KD, M100.G_NAME, M010.EMPL_NAME, AA.G_CODE, ZTBPOTable.PO_QTY, ZTBPOTable.PROD_PRICE, AA.TotalDelivered as TOTAL_DELIVERED, (ZTBPOTable.PO_QTY-AA.TotalDelivered) As PO_BALANCE,(ZTBPOTable.PO_QTY*ZTBPOTable.PROD_PRICE) As PO_AMOUNT , (AA.TotalDelivered*ZTBPOTable.PROD_PRICE) As DELIVERED_AMOUNT, ((ZTBPOTable.PO_QTY-AA.TotalDelivered)*ZTBPOTable.PROD_PRICE) As BALANCE_AMOUNT,DATEPART( MONTH, PO_DATE) AS POMONTH, DATEPART( ISOWK, PO_DATE) AS POWEEKNUM, 				CASE 					WHEN (ZTBPOTable.RD_DATE < GETDATE()-1) AND ((ZTBPOTable.PO_QTY-AA.TotalDelivered) <>0) THEN 'OVER'    					ELSE 'OK' 				END AS OVERDUE, ZTBPOTable.REMARK, ZTBPOTable.PO_ID 				FROM 				(SELECT ZTBPOTable.EMPL_NO, ZTBPOTable.CUST_CD, ZTBPOTable.G_CODE, ZTBPOTable.PO_NO, isnull(SUM(ZTBDelivery.DELIVERY_QTY),0) AS TotalDelivered 				FROM 				fn_PODATE(@ngaythang)  AS ZTBPOTable 				LEFT JOIN 				fn_DELIVERYTODATE(@ngaythang) AS  ZTBDelivery 				ON (ZTBDelivery.CTR_CD = ZTBPOTable.CTR_CD AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) 				GROUP BY ZTBPOTable.CTR_CD,ZTBPOTable.EMPL_NO,ZTBPOTable.G_CODE,ZTBPOTable.CUST_CD,ZTBPOTable.PO_NO) AS AA 				LEFT JOIN M010 				ON (M010.EMPL_NO = AA.EMPL_NO) 				LEFT JOIN M100 				ON (M100.G_CODE = AA.G_CODE) 				LEFT JOIN ZTBPOTable 				ON (AA.CUST_CD = ZTBPOTable.CUST_CD AND AA.G_CODE = ZTBPOTable.G_CODE AND AA.PO_NO = ZTBPOTable.PO_NO) 				JOIN M110 				ON (M110.CUST_CD = AA.CUST_CD) 	 	  			) AS P 		) AS j 		PIVOT (SUM(j.PO_BALANCE) FOR j.PROD_TYPE IN ([TSP],[LABEL],[UV],[OLED],[TAPE],[RIBBON],[SPT])) AS PV 	) AS AA 	JOIN 	(   	SELECT A.POWEEKNUM, SUM(A.PO_BALANCE) As TOTAL_PO_BALANCE 	FROM 	(   		SELECT AA.PO_NO,  M100.PROD_TYPE, M100.PROD_MAIN_MATERIAL, ZTBPOTable.PO_DATE,ZTBPOTable.RD_DATE, M110.CUST_NAME_KD, M100.G_NAME, M010.EMPL_NAME, AA.G_CODE, ZTBPOTable.PO_QTY, ZTBPOTable.PROD_PRICE, AA.TotalDelivered as TOTAL_DELIVERED, (ZTBPOTable.PO_QTY-AA.TotalDelivered) As PO_BALANCE,(ZTBPOTable.PO_QTY*ZTBPOTable.PROD_PRICE) As PO_AMOUNT , (AA.TotalDelivered*ZTBPOTable.PROD_PRICE) As DELIVERED_AMOUNT, ((ZTBPOTable.PO_QTY-AA.TotalDelivered)*ZTBPOTable.PROD_PRICE) As BALANCE_AMOUNT,DATEPART( MONTH, PO_DATE) AS POMONTH, DATEPART( ISOWK, PO_DATE) AS POWEEKNUM, 				CASE 					WHEN (ZTBPOTable.RD_DATE < GETDATE()-1) AND ((ZTBPOTable.PO_QTY-AA.TotalDelivered) <>0) THEN 'OVER'    					ELSE 'OK' 				END AS OVERDUE, ZTBPOTable.REMARK, ZTBPOTable.PO_ID 				FROM 				(SELECT ZTBPOTable.EMPL_NO, ZTBPOTable.CUST_CD, ZTBPOTable.G_CODE, ZTBPOTable.PO_NO, isnull(SUM(ZTBDelivery.DELIVERY_QTY),0) AS TotalDelivered 				FROM 				fn_PODATE(@ngaythang)  AS ZTBPOTable 				LEFT JOIN 				fn_DELIVERYTODATE(@ngaythang) AS  ZTBDelivery 				ON (ZTBDelivery.CTR_CD = ZTBPOTable.CTR_CD AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) 				GROUP BY ZTBPOTable.CTR_CD,ZTBPOTable.EMPL_NO,ZTBPOTable.G_CODE,ZTBPOTable.CUST_CD,ZTBPOTable.PO_NO) AS AA 				LEFT JOIN M010 				ON (M010.EMPL_NO = AA.EMPL_NO) 				LEFT JOIN M100 				ON (M100.G_CODE = AA.G_CODE) 				LEFT JOIN ZTBPOTable 				ON (AA.CUST_CD = ZTBPOTable.CUST_CD AND AA.G_CODE = ZTBPOTable.G_CODE AND AA.PO_NO = ZTBPOTable.PO_NO) 				JOIN M110 				ON (M110.CUST_CD = AA.CUST_CD)   	) AS A 	GROUP BY POWEEKNUM 	) AS BB 	ON AA.POWEEKNUM = BB.POWEEKNUM ) AS XX";
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable report_POBalanceByCustomer(string balanceDate)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "USE CMS_VINA DECLARE @ngaythang AS DATE SET @ngaythang = '"+balanceDate + "'  SELECT YEAR(@ngaythang) AS BALANCE_YEAR, DATEPART( ISOWK, @ngaythang) AS BALANCE_WEEKNUM ,@ngaythang AS BALANCE_DATE, SUM(TOTAL_PO_BALANCE) AS TOTAL_PO_BALANCE, SUM(SEV) AS SEV, SUM(SEVT) AS SEVT, SUM(SAMSUNG_ASIA) AS SAMSUNG_ASIA, (SUM(TOTAL_PO_BALANCE) -  SUM(SEV) -SUM(SEVT) -SUM(SAMSUNG_ASIA)) AS OTHERS FROM ( 	SELECT AA.POWEEKNUM, BB.TOTAL_PO_BALANCE, AA.SEV, AA.SEVT, AA.SAMSUNG_ASIA,  (BB.TOTAL_PO_BALANCE-AA.SEV-AA.SEVT-AA.SAMSUNG_ASIA) AS OTHERS FROM 	( 		SELECT  POWEEKNUM,   isnull(PV.[SEV],0) As SEV, isnull(PV.[SEVT],0) As SEVT,isnull(PV.[SAMSUNG-ASIA],0) As SAMSUNG_ASIA 		FROM 		( 		SELECT P.PO_BALANCE, P.CUST_NAME_KD,  P.POWEEKNUM FROM 		( 	 			SELECT AA.PO_NO,  M100.PROD_TYPE, M100.PROD_MAIN_MATERIAL, ZTBPOTable.PO_DATE,ZTBPOTable.RD_DATE, M110.CUST_NAME_KD, M100.G_NAME, M010.EMPL_NAME, AA.G_CODE, ZTBPOTable.PO_QTY, ZTBPOTable.PROD_PRICE, AA.TotalDelivered as TOTAL_DELIVERED, (ZTBPOTable.PO_QTY-AA.TotalDelivered) As PO_BALANCE,(ZTBPOTable.PO_QTY*ZTBPOTable.PROD_PRICE) As PO_AMOUNT , (AA.TotalDelivered*ZTBPOTable.PROD_PRICE) As DELIVERED_AMOUNT, ((ZTBPOTable.PO_QTY-AA.TotalDelivered)*ZTBPOTable.PROD_PRICE) As BALANCE_AMOUNT,DATEPART( MONTH, PO_DATE) AS POMONTH, DATEPART( ISOWK, PO_DATE) AS POWEEKNUM, 			CASE 				WHEN (ZTBPOTable.RD_DATE < GETDATE()-1) AND ((ZTBPOTable.PO_QTY-AA.TotalDelivered) <>0) THEN 'OVER'    				ELSE 'OK' 			END AS OVERDUE, ZTBPOTable.REMARK, ZTBPOTable.PO_ID 			FROM 			(SELECT ZTBPOTable.EMPL_NO, ZTBPOTable.CUST_CD, ZTBPOTable.G_CODE, ZTBPOTable.PO_NO, isnull(SUM(ZTBDelivery.DELIVERY_QTY),0) AS TotalDelivered 			FROM 			fn_PODATE(@ngaythang)  AS ZTBPOTable 			LEFT JOIN 			fn_DELIVERYTODATE(@ngaythang) AS  ZTBDelivery 			ON (ZTBDelivery.CTR_CD = ZTBPOTable.CTR_CD AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) 			GROUP BY ZTBPOTable.CTR_CD,ZTBPOTable.EMPL_NO,ZTBPOTable.G_CODE,ZTBPOTable.CUST_CD,ZTBPOTable.PO_NO) AS AA 			LEFT JOIN M010 			ON (M010.EMPL_NO = AA.EMPL_NO) 			LEFT JOIN M100 			ON (M100.G_CODE = AA.G_CODE) 			LEFT JOIN ZTBPOTable 			ON (AA.CUST_CD = ZTBPOTable.CUST_CD AND AA.G_CODE = ZTBPOTable.G_CODE AND AA.PO_NO = ZTBPOTable.PO_NO) 			JOIN M110 			ON (M110.CUST_CD = AA.CUST_CD) 	 	  		) AS P 		) AS j 		PIVOT (SUM(j.PO_BALANCE) FOR j.CUST_NAME_KD IN ([SEV],[SEVT],[SAMSUNG-ASIA])) AS PV 	) AS AA 	JOIN 	(    	SELECT A.POWEEKNUM, SUM(A.PO_BALANCE) As TOTAL_PO_BALANCE 	FROM 	(   		SELECT AA.PO_NO,  M100.PROD_TYPE, M100.PROD_MAIN_MATERIAL, ZTBPOTable.PO_DATE,ZTBPOTable.RD_DATE, M110.CUST_NAME_KD, M100.G_NAME, M010.EMPL_NAME, AA.G_CODE, ZTBPOTable.PO_QTY, ZTBPOTable.PROD_PRICE, AA.TotalDelivered as TOTAL_DELIVERED, (ZTBPOTable.PO_QTY-AA.TotalDelivered) As PO_BALANCE,(ZTBPOTable.PO_QTY*ZTBPOTable.PROD_PRICE) As PO_AMOUNT , (AA.TotalDelivered*ZTBPOTable.PROD_PRICE) As DELIVERED_AMOUNT, ((ZTBPOTable.PO_QTY-AA.TotalDelivered)*ZTBPOTable.PROD_PRICE) As BALANCE_AMOUNT,DATEPART( MONTH, PO_DATE) AS POMONTH, DATEPART( ISOWK, PO_DATE) AS POWEEKNUM, 				CASE 					WHEN (ZTBPOTable.RD_DATE < GETDATE()-1) AND ((ZTBPOTable.PO_QTY-AA.TotalDelivered) <>0) THEN 'OVER'    					ELSE 'OK' 				END AS OVERDUE, ZTBPOTable.REMARK, ZTBPOTable.PO_ID 				FROM 				(SELECT ZTBPOTable.EMPL_NO, ZTBPOTable.CUST_CD, ZTBPOTable.G_CODE, ZTBPOTable.PO_NO, isnull(SUM(ZTBDelivery.DELIVERY_QTY),0) AS TotalDelivered 				FROM 				fn_PODATE(@ngaythang)  AS ZTBPOTable 				LEFT JOIN 				fn_DELIVERYTODATE(@ngaythang) AS  ZTBDelivery 				ON (ZTBDelivery.CTR_CD = ZTBPOTable.CTR_CD AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) 				GROUP BY ZTBPOTable.CTR_CD,ZTBPOTable.EMPL_NO,ZTBPOTable.G_CODE,ZTBPOTable.CUST_CD,ZTBPOTable.PO_NO) AS AA 				LEFT JOIN M010 				ON (M010.EMPL_NO = AA.EMPL_NO) 				LEFT JOIN M100 				ON (M100.G_CODE = AA.G_CODE) 				LEFT JOIN ZTBPOTable 				ON (AA.CUST_CD = ZTBPOTable.CUST_CD AND AA.G_CODE = ZTBPOTable.G_CODE AND AA.PO_NO = ZTBPOTable.PO_NO) 				JOIN M110 				ON (M110.CUST_CD = AA.CUST_CD)   	) AS A 	GROUP BY POWEEKNUM 	) AS BB 	ON AA.POWEEKNUM = BB.POWEEKNUM ) AS XX";
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable report_WeeklyPOByType(string potoday)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT  PO_YEAR, POWEEKNUM, '"+ potoday + "' AS TODAYDATE, (isnull(PV.[TSP],0)+isnull(PV.[LABEL],0)+isnull(PV.[UV],0)+isnull(PV.[TAPE],0) + isnull(PV.[SPT],0)+ isnull(PV.[OLED],0) + isnull(PV.[NODATA],0) + isnull(PV.[RIBBON],0)) As TOTAL_PO_QTY, isnull(PV.[TSP],0) As TSP, isnull(PV.[LABEL],0) As LABEL,isnull(PV.[UV],0) As UV, isnull(PV.[OLED],0) As OLED,isnull(PV.[TAPE],0) As TAPE,  isnull(PV.[RIBBON],0) As RIBBON, isnull(PV.[SPT],0) As SPT, isnull(PV.[NODATA],0) As OTHERS FROM ( SELECT P.PO_QTY, P.PROD_TYPE,  P.POWEEKNUM, P.PO_YEAR FROM (   SELECT AA.PO_NO,  M100.PROD_TYPE, M100.PROD_MAIN_MATERIAL, ZTBPOTable.PO_DATE,ZTBPOTable.RD_DATE, M110.CUST_NAME_KD, M100.G_NAME, M010.EMPL_NAME, AA.G_CODE, ZTBPOTable.PO_QTY, ZTBPOTable.PROD_PRICE, AA.TotalDelivered as TOTAL_DELIVERED, (ZTBPOTable.PO_QTY-AA.TotalDelivered) As PO_BALANCE,(ZTBPOTable.PO_QTY*ZTBPOTable.PROD_PRICE) As PO_AMOUNT , (AA.TotalDelivered*ZTBPOTable.PROD_PRICE) As DELIVERED_AMOUNT, ((ZTBPOTable.PO_QTY-AA.TotalDelivered)*ZTBPOTable.PROD_PRICE) As BALANCE_AMOUNT,DATEPART( MONTH, PO_DATE) AS POMONTH, DATEPART( ISOWK, PO_DATE) AS POWEEKNUM, YEAR(PO_DATE) As PO_YEAR, CASE     WHEN (ZTBPOTable.RD_DATE < GETDATE()-1) AND ((ZTBPOTable.PO_QTY-AA.TotalDelivered) <>0) THEN 'OVER'        ELSE 'OK' END AS OVERDUE, ZTBPOTable.REMARK, ZTBPOTable.PO_ID FROM (SELECT ZTBPOTable.EMPL_NO, ZTBPOTable.CUST_CD, ZTBPOTable.G_CODE, ZTBPOTable.PO_NO, isnull(SUM(ZTBDelivery.DELIVERY_QTY),0) AS TotalDelivered FROM ZTBPOTable  LEFT JOIN ZTBDelivery ON (ZTBDelivery.CTR_CD = ZTBPOTable.CTR_CD AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) GROUP BY ZTBPOTable.CTR_CD,ZTBPOTable.EMPL_NO,ZTBPOTable.G_CODE,ZTBPOTable.CUST_CD,ZTBPOTable.PO_NO) AS AA LEFT JOIN M010 ON (M010.EMPL_NO = AA.EMPL_NO) LEFT JOIN M100 ON (M100.G_CODE = AA.G_CODE) LEFT JOIN ZTBPOTable ON (AA.CUST_CD = ZTBPOTable.CUST_CD AND AA.G_CODE = ZTBPOTable.G_CODE AND AA.PO_NO = ZTBPOTable.PO_NO) JOIN M110 ON (M110.CUST_CD = AA.CUST_CD) /*WHERE CUST_NAME_KD='SEVT'*/   ) AS P ) AS j PIVOT (SUM(j.PO_QTY) FOR j.PROD_TYPE IN ([TSP],[LABEL],[UV],[OLED],[TAPE],[SPT],[RIBBON],[NODATA])) AS PV ORDER BY POWEEKNUM DESC";
            result = config.GetData(strQuery);
            return result;
        }

        public DataTable report_WeeklyPOByTypeALL2SS(string startdate, string enddate)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "DECLARE @startdate DATE DECLARE @enddate DATE SET @startdate ='" + startdate + "' SET @enddate = '" + enddate + "' SELECT KK.START_DATE, KK.END_DATE, NN.TOTAL_PO_QTY, KK.TSP, KK.LABEL, KK.UV,KK.OLED, KK.TAPE, KK.RIBBON, KK.SPT, (NN.TOTAL_PO_QTY-KK.TSP- KK.LABEL- KK.UV-KK.OLED- KK.TAPE- KK.RIBBON- KK.SPT) AS OTHERS FROM ( SELECT @startdate AS START_DATE, @enddate as END_DATE, isnull(PV.[TSP],0) As TSP, isnull(PV.[LABEL],0) As LABEL,isnull(PV.[UV],0) As UV, isnull(PV.[OLED],0) As OLED,isnull(PV.[TAPE],0) As TAPE, isnull(PV.[SPT],0) As SPT, isnull(PV.[RIBBON],0) As RIBBON  FROM ( SELECT XX.PROD_TYPE, XX.PO_QTY FROM (SELECT AA.PO_NO,  M100.PROD_TYPE, M100.PROD_MAIN_MATERIAL, ZTBPOTable.PO_DATE,ZTBPOTable.RD_DATE, M110.CUST_NAME_KD, M100.G_NAME, M010.EMPL_NAME, AA.G_CODE, ZTBPOTable.PO_QTY, ZTBPOTable.PROD_PRICE, AA.TotalDelivered as TOTAL_DELIVERED, (ZTBPOTable.PO_QTY-AA.TotalDelivered) As PO_BALANCE,(ZTBPOTable.PO_QTY*ZTBPOTable.PROD_PRICE) As PO_AMOUNT , (AA.TotalDelivered*ZTBPOTable.PROD_PRICE) As DELIVERED_AMOUNT, ((ZTBPOTable.PO_QTY-AA.TotalDelivered)*ZTBPOTable.PROD_PRICE) As BALANCE_AMOUNT,DATEPART( MONTH, PO_DATE) AS POMONTH, DATEPART( ISOWK, PO_DATE) AS POWEEKNUM, YEAR(PO_DATE) As PO_YEAR, 	CASE 		WHEN (ZTBPOTable.RD_DATE < GETDATE()-1) AND ((ZTBPOTable.PO_QTY-AA.TotalDelivered) <>0) THEN 'OVER'    		ELSE 'OK' 	END AS OVERDUE, ZTBPOTable.REMARK, ZTBPOTable.PO_ID 	FROM 	(SELECT ZTBPOTable.EMPL_NO, ZTBPOTable.CUST_CD, ZTBPOTable.G_CODE, ZTBPOTable.PO_NO, isnull(SUM(ZTBDelivery.DELIVERY_QTY),0) AS TotalDelivered 	FROM ZTBPOTable  LEFT JOIN ZTBDelivery 	ON (ZTBDelivery.CTR_CD = ZTBPOTable.CTR_CD AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) 	GROUP BY ZTBPOTable.CTR_CD,ZTBPOTable.EMPL_NO,ZTBPOTable.G_CODE,ZTBPOTable.CUST_CD,ZTBPOTable.PO_NO) AS AA 	LEFT JOIN M010 	ON (M010.EMPL_NO = AA.EMPL_NO) 	LEFT JOIN M100 	ON (M100.G_CODE = AA.G_CODE) 	LEFT JOIN ZTBPOTable 	ON (AA.CUST_CD = ZTBPOTable.CUST_CD AND AA.G_CODE = ZTBPOTable.G_CODE AND AA.PO_NO = ZTBPOTable.PO_NO) 	JOIN M110 	ON (M110.CUST_CD = AA.CUST_CD) 	WHERE PO_DATE BETWEEN @startdate AND @enddate AND CUST_NAME_KD IN ('SEV','SEVT','SAMSUNG-ASIA') ) AS XX ) AS YY PIVOT (SUM(YY.PO_QTY) FOR YY.PROD_TYPE IN ([TSP],[LABEL],[UV],[OLED],[TAPE],[SPT],[RIBBON])) AS PV ) AS KK JOIN ( SELECT @startdate AS START_DATE, @enddate AS END_DATE, SUM(MM.PO_QTY) AS TOTAL_PO_QTY FROM (SELECT AA.PO_NO,  M100.PROD_TYPE, M100.PROD_MAIN_MATERIAL, ZTBPOTable.PO_DATE,ZTBPOTable.RD_DATE, M110.CUST_NAME_KD, M100.G_NAME, M010.EMPL_NAME, AA.G_CODE, ZTBPOTable.PO_QTY, ZTBPOTable.PROD_PRICE, AA.TotalDelivered as TOTAL_DELIVERED, (ZTBPOTable.PO_QTY-AA.TotalDelivered) As PO_BALANCE,(ZTBPOTable.PO_QTY*ZTBPOTable.PROD_PRICE) As PO_AMOUNT , (AA.TotalDelivered*ZTBPOTable.PROD_PRICE) As DELIVERED_AMOUNT, ((ZTBPOTable.PO_QTY-AA.TotalDelivered)*ZTBPOTable.PROD_PRICE) As BALANCE_AMOUNT,DATEPART( MONTH, PO_DATE) AS POMONTH, DATEPART( ISOWK, PO_DATE) AS POWEEKNUM, YEAR(PO_DATE) As PO_YEAR, 	CASE 		WHEN (ZTBPOTable.RD_DATE < GETDATE()-1) AND ((ZTBPOTable.PO_QTY-AA.TotalDelivered) <>0) THEN 'OVER'    		ELSE 'OK' 	END AS OVERDUE, ZTBPOTable.REMARK, ZTBPOTable.PO_ID 	FROM 	(SELECT ZTBPOTable.EMPL_NO, ZTBPOTable.CUST_CD, ZTBPOTable.G_CODE, ZTBPOTable.PO_NO, isnull(SUM(ZTBDelivery.DELIVERY_QTY),0) AS TotalDelivered 	FROM ZTBPOTable  LEFT JOIN ZTBDelivery 	ON (ZTBDelivery.CTR_CD = ZTBPOTable.CTR_CD AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) 	GROUP BY ZTBPOTable.CTR_CD,ZTBPOTable.EMPL_NO,ZTBPOTable.G_CODE,ZTBPOTable.CUST_CD,ZTBPOTable.PO_NO) AS AA 	LEFT JOIN M010 	ON (M010.EMPL_NO = AA.EMPL_NO) 	LEFT JOIN M100 	ON (M100.G_CODE = AA.G_CODE) 	LEFT JOIN ZTBPOTable 	ON (AA.CUST_CD = ZTBPOTable.CUST_CD AND AA.G_CODE = ZTBPOTable.G_CODE AND AA.PO_NO = ZTBPOTable.PO_NO) 	JOIN M110 	ON (M110.CUST_CD = AA.CUST_CD) 	WHERE PO_DATE BETWEEN @startdate AND @enddate AND CUST_NAME_KD IN ('SEV','SEVT','SAMSUNG-ASIA') ) AS MM ) AS NN ON (KK.START_DATE = NN.START_DATE AND KK.END_DATE = NN.END_DATE)";
            result = config.GetData(strQuery);
            return result;

        }


        public DataTable report_WeeklyPOByTypeALL2(string startdate, string enddate)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "DECLARE @startdate DATE DECLARE @enddate DATE SET @startdate ='" + startdate + "' SET @enddate = '" + enddate + "' SELECT KK.START_DATE, KK.END_DATE, NN.TOTAL_PO_QTY, KK.TSP, KK.LABEL, KK.UV,KK.OLED, KK.TAPE, KK.RIBBON, KK.SPT, (NN.TOTAL_PO_QTY-KK.TSP- KK.LABEL- KK.UV-KK.OLED- KK.TAPE- KK.RIBBON- KK.SPT) AS OTHERS FROM ( SELECT @startdate AS START_DATE, @enddate as END_DATE, isnull(PV.[TSP],0) As TSP, isnull(PV.[LABEL],0) As LABEL,isnull(PV.[UV],0) As UV, isnull(PV.[OLED],0) As OLED,isnull(PV.[TAPE],0) As TAPE, isnull(PV.[SPT],0) As SPT, isnull(PV.[RIBBON],0) As RIBBON  FROM ( SELECT XX.PROD_TYPE, XX.PO_QTY FROM (SELECT AA.PO_NO,  M100.PROD_TYPE, M100.PROD_MAIN_MATERIAL, ZTBPOTable.PO_DATE,ZTBPOTable.RD_DATE, M110.CUST_NAME_KD, M100.G_NAME, M010.EMPL_NAME, AA.G_CODE, ZTBPOTable.PO_QTY, ZTBPOTable.PROD_PRICE, AA.TotalDelivered as TOTAL_DELIVERED, (ZTBPOTable.PO_QTY-AA.TotalDelivered) As PO_BALANCE,(ZTBPOTable.PO_QTY*ZTBPOTable.PROD_PRICE) As PO_AMOUNT , (AA.TotalDelivered*ZTBPOTable.PROD_PRICE) As DELIVERED_AMOUNT, ((ZTBPOTable.PO_QTY-AA.TotalDelivered)*ZTBPOTable.PROD_PRICE) As BALANCE_AMOUNT,DATEPART( MONTH, PO_DATE) AS POMONTH, DATEPART( ISOWK, PO_DATE) AS POWEEKNUM, YEAR(PO_DATE) As PO_YEAR, 	CASE 		WHEN (ZTBPOTable.RD_DATE < GETDATE()-1) AND ((ZTBPOTable.PO_QTY-AA.TotalDelivered) <>0) THEN 'OVER'    		ELSE 'OK' 	END AS OVERDUE, ZTBPOTable.REMARK, ZTBPOTable.PO_ID 	FROM 	(SELECT ZTBPOTable.EMPL_NO, ZTBPOTable.CUST_CD, ZTBPOTable.G_CODE, ZTBPOTable.PO_NO, isnull(SUM(ZTBDelivery.DELIVERY_QTY),0) AS TotalDelivered 	FROM ZTBPOTable  LEFT JOIN ZTBDelivery 	ON (ZTBDelivery.CTR_CD = ZTBPOTable.CTR_CD AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) 	GROUP BY ZTBPOTable.CTR_CD,ZTBPOTable.EMPL_NO,ZTBPOTable.G_CODE,ZTBPOTable.CUST_CD,ZTBPOTable.PO_NO) AS AA 	LEFT JOIN M010 	ON (M010.EMPL_NO = AA.EMPL_NO) 	LEFT JOIN M100 	ON (M100.G_CODE = AA.G_CODE) 	LEFT JOIN ZTBPOTable 	ON (AA.CUST_CD = ZTBPOTable.CUST_CD AND AA.G_CODE = ZTBPOTable.G_CODE AND AA.PO_NO = ZTBPOTable.PO_NO) 	JOIN M110 	ON (M110.CUST_CD = AA.CUST_CD) 	WHERE PO_DATE BETWEEN @startdate AND @enddate ) AS XX ) AS YY PIVOT (SUM(YY.PO_QTY) FOR YY.PROD_TYPE IN ([TSP],[LABEL],[UV],[OLED],[TAPE],[SPT],[RIBBON])) AS PV ) AS KK JOIN ( SELECT @startdate AS START_DATE, @enddate AS END_DATE, SUM(MM.PO_QTY) AS TOTAL_PO_QTY FROM (SELECT AA.PO_NO,  M100.PROD_TYPE, M100.PROD_MAIN_MATERIAL, ZTBPOTable.PO_DATE,ZTBPOTable.RD_DATE, M110.CUST_NAME_KD, M100.G_NAME, M010.EMPL_NAME, AA.G_CODE, ZTBPOTable.PO_QTY, ZTBPOTable.PROD_PRICE, AA.TotalDelivered as TOTAL_DELIVERED, (ZTBPOTable.PO_QTY-AA.TotalDelivered) As PO_BALANCE,(ZTBPOTable.PO_QTY*ZTBPOTable.PROD_PRICE) As PO_AMOUNT , (AA.TotalDelivered*ZTBPOTable.PROD_PRICE) As DELIVERED_AMOUNT, ((ZTBPOTable.PO_QTY-AA.TotalDelivered)*ZTBPOTable.PROD_PRICE) As BALANCE_AMOUNT,DATEPART( MONTH, PO_DATE) AS POMONTH, DATEPART( ISOWK, PO_DATE) AS POWEEKNUM, YEAR(PO_DATE) As PO_YEAR, 	CASE 		WHEN (ZTBPOTable.RD_DATE < GETDATE()-1) AND ((ZTBPOTable.PO_QTY-AA.TotalDelivered) <>0) THEN 'OVER'    		ELSE 'OK' 	END AS OVERDUE, ZTBPOTable.REMARK, ZTBPOTable.PO_ID 	FROM 	(SELECT ZTBPOTable.EMPL_NO, ZTBPOTable.CUST_CD, ZTBPOTable.G_CODE, ZTBPOTable.PO_NO, isnull(SUM(ZTBDelivery.DELIVERY_QTY),0) AS TotalDelivered 	FROM ZTBPOTable  LEFT JOIN ZTBDelivery 	ON (ZTBDelivery.CTR_CD = ZTBPOTable.CTR_CD AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) 	GROUP BY ZTBPOTable.CTR_CD,ZTBPOTable.EMPL_NO,ZTBPOTable.G_CODE,ZTBPOTable.CUST_CD,ZTBPOTable.PO_NO) AS AA 	LEFT JOIN M010 	ON (M010.EMPL_NO = AA.EMPL_NO) 	LEFT JOIN M100 	ON (M100.G_CODE = AA.G_CODE) 	LEFT JOIN ZTBPOTable 	ON (AA.CUST_CD = ZTBPOTable.CUST_CD AND AA.G_CODE = ZTBPOTable.G_CODE AND AA.PO_NO = ZTBPOTable.PO_NO) 	JOIN M110 	ON (M110.CUST_CD = AA.CUST_CD) 	WHERE PO_DATE BETWEEN @startdate AND @enddate ) AS MM ) AS NN ON (KK.START_DATE = NN.START_DATE AND KK.END_DATE = NN.END_DATE)";
            result = config.GetData(strQuery);
            return result;

        }

    


        public DataTable report_WeeklyPOByCustomer2(string startdate, string enddate)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "DECLARE @startdate DATE DECLARE @enddate DATE SET @startdate ='" + startdate + "' SET @enddate = '" + enddate + "' SELECT KK.START_DATE, KK.END_DATE, NN.TOTAL_PO_QTY, KK.SEV, KK.SEVT, KK.SAMSUNG_ASIA, (NN.TOTAL_PO_QTY-KK.SEV- KK.SEVT- KK.SAMSUNG_ASIA) AS OTHERS FROM ( SELECT @startdate AS START_DATE, @enddate as END_DATE, isnull([SEV],0) AS SEV,isnull([SEVT],0) AS SEVT,isnull([SAMSUNG-ASIA],0) AS SAMSUNG_ASIA  FROM ( SELECT XX.CUST_NAME_KD, XX.PO_QTY FROM ( SELECT AA.PO_NO,  M100.PROD_TYPE, M100.PROD_MAIN_MATERIAL, ZTBPOTable.PO_DATE,ZTBPOTable.RD_DATE, M110.CUST_NAME_KD, M100.G_NAME, M010.EMPL_NAME, AA.G_CODE, ZTBPOTable.PO_QTY, ZTBPOTable.PROD_PRICE, AA.TotalDelivered as TOTAL_DELIVERED, (ZTBPOTable.PO_QTY-AA.TotalDelivered) As PO_BALANCE,(ZTBPOTable.PO_QTY*ZTBPOTable.PROD_PRICE) As PO_AMOUNT , (AA.TotalDelivered*ZTBPOTable.PROD_PRICE) As DELIVERED_AMOUNT, ((ZTBPOTable.PO_QTY-AA.TotalDelivered)*ZTBPOTable.PROD_PRICE) As BALANCE_AMOUNT,DATEPART( MONTH, PO_DATE) AS POMONTH, DATEPART( ISOWK, PO_DATE) AS POWEEKNUM, YEAR(PO_DATE) As PO_YEAR, 	CASE 		WHEN (ZTBPOTable.RD_DATE < GETDATE()-1) AND ((ZTBPOTable.PO_QTY-AA.TotalDelivered) <>0) THEN 'OVER'    		ELSE 'OK' 	END AS OVERDUE, ZTBPOTable.REMARK, ZTBPOTable.PO_ID 	FROM 	(SELECT ZTBPOTable.EMPL_NO, ZTBPOTable.CUST_CD, ZTBPOTable.G_CODE, ZTBPOTable.PO_NO, isnull(SUM(ZTBDelivery.DELIVERY_QTY),0) AS TotalDelivered 	FROM ZTBPOTable  LEFT JOIN ZTBDelivery 	ON (ZTBDelivery.CTR_CD = ZTBPOTable.CTR_CD AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) 	GROUP BY ZTBPOTable.CTR_CD,ZTBPOTable.EMPL_NO,ZTBPOTable.G_CODE,ZTBPOTable.CUST_CD,ZTBPOTable.PO_NO) AS AA 	LEFT JOIN M010 	ON (M010.EMPL_NO = AA.EMPL_NO) 	LEFT JOIN M100 	ON (M100.G_CODE = AA.G_CODE) 	LEFT JOIN ZTBPOTable 	ON (AA.CUST_CD = ZTBPOTable.CUST_CD AND AA.G_CODE = ZTBPOTable.G_CODE AND AA.PO_NO = ZTBPOTable.PO_NO) 	JOIN M110 	ON (M110.CUST_CD = AA.CUST_CD) 	WHERE PO_DATE BETWEEN @startdate AND @enddate ) AS XX ) AS YY PIVOT (SUM(YY.PO_QTY) FOR YY.CUST_NAME_KD IN ([SEV],[SEVT],[SAMSUNG-ASIA])) AS PV ) AS KK JOIN ( SELECT @startdate AS START_DATE, @enddate AS END_DATE, SUM(MM.PO_QTY) AS TOTAL_PO_QTY FROM (SELECT AA.PO_NO,  M100.PROD_TYPE, M100.PROD_MAIN_MATERIAL, ZTBPOTable.PO_DATE,ZTBPOTable.RD_DATE, M110.CUST_NAME_KD, M100.G_NAME, M010.EMPL_NAME, AA.G_CODE, ZTBPOTable.PO_QTY, ZTBPOTable.PROD_PRICE, AA.TotalDelivered as TOTAL_DELIVERED, (ZTBPOTable.PO_QTY-AA.TotalDelivered) As PO_BALANCE,(ZTBPOTable.PO_QTY*ZTBPOTable.PROD_PRICE) As PO_AMOUNT , (AA.TotalDelivered*ZTBPOTable.PROD_PRICE) As DELIVERED_AMOUNT, ((ZTBPOTable.PO_QTY-AA.TotalDelivered)*ZTBPOTable.PROD_PRICE) As BALANCE_AMOUNT,DATEPART( MONTH, PO_DATE) AS POMONTH, DATEPART( ISOWK, PO_DATE) AS POWEEKNUM, YEAR(PO_DATE) As PO_YEAR, 	CASE 		WHEN (ZTBPOTable.RD_DATE < GETDATE()-1) AND ((ZTBPOTable.PO_QTY-AA.TotalDelivered) <>0) THEN 'OVER'    		ELSE 'OK' 	END AS OVERDUE, ZTBPOTable.REMARK, ZTBPOTable.PO_ID 	FROM 	(SELECT ZTBPOTable.EMPL_NO, ZTBPOTable.CUST_CD, ZTBPOTable.G_CODE, ZTBPOTable.PO_NO, isnull(SUM(ZTBDelivery.DELIVERY_QTY),0) AS TotalDelivered 	FROM ZTBPOTable  LEFT JOIN ZTBDelivery 	ON (ZTBDelivery.CTR_CD = ZTBPOTable.CTR_CD AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) 	GROUP BY ZTBPOTable.CTR_CD,ZTBPOTable.EMPL_NO,ZTBPOTable.G_CODE,ZTBPOTable.CUST_CD,ZTBPOTable.PO_NO) AS AA 	LEFT JOIN M010 	ON (M010.EMPL_NO = AA.EMPL_NO) 	LEFT JOIN M100 	ON (M100.G_CODE = AA.G_CODE) 	LEFT JOIN ZTBPOTable 	ON (AA.CUST_CD = ZTBPOTable.CUST_CD AND AA.G_CODE = ZTBPOTable.G_CODE AND AA.PO_NO = ZTBPOTable.PO_NO) 	JOIN M110 	ON (M110.CUST_CD = AA.CUST_CD) 	WHERE PO_DATE BETWEEN @startdate AND @enddate ) AS MM ) AS NN ON (KK.START_DATE = NN.START_DATE AND KK.END_DATE = NN.END_DATE)";
            result = config.GetData(strQuery);
            return result;

        }


        public DataTable report_WeeklyPOByTypeSS2(string startdate, string enddate)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "DECLARE @startdate DATE DECLARE @enddate DATE SET @startdate ='" + startdate + "' SET @enddate = '" + enddate + "' SELECT KK.START_DATE, KK.END_DATE, NN.TOTAL_PO_QTY, KK.SEV, KK.SEVT, KK.SAMSUNG_ASIA, (NN.TOTAL_PO_QTY-KK.SEV- KK.SEVT- KK.SAMSUNG_ASIA) AS OTHERS FROM ( SELECT @startdate AS START_DATE, @enddate as END_DATE, isnull([SEV],0) AS SEV,isnull([SEVT],0) AS SEVT,isnull([SAMSUNG-ASIA],0) AS SAMSUNG_ASIA  FROM ( SELECT XX.CUST_NAME_KD, XX.PO_QTY FROM ( SELECT AA.PO_NO,  M100.PROD_TYPE, M100.PROD_MAIN_MATERIAL, ZTBPOTable.PO_DATE,ZTBPOTable.RD_DATE, M110.CUST_NAME_KD, M100.G_NAME, M010.EMPL_NAME, AA.G_CODE, ZTBPOTable.PO_QTY, ZTBPOTable.PROD_PRICE, AA.TotalDelivered as TOTAL_DELIVERED, (ZTBPOTable.PO_QTY-AA.TotalDelivered) As PO_BALANCE,(ZTBPOTable.PO_QTY*ZTBPOTable.PROD_PRICE) As PO_AMOUNT , (AA.TotalDelivered*ZTBPOTable.PROD_PRICE) As DELIVERED_AMOUNT, ((ZTBPOTable.PO_QTY-AA.TotalDelivered)*ZTBPOTable.PROD_PRICE) As BALANCE_AMOUNT,DATEPART( MONTH, PO_DATE) AS POMONTH, DATEPART( ISOWK, PO_DATE) AS POWEEKNUM, YEAR(PO_DATE) As PO_YEAR, 	CASE 		WHEN (ZTBPOTable.RD_DATE < GETDATE()-1) AND ((ZTBPOTable.PO_QTY-AA.TotalDelivered) <>0) THEN 'OVER'    		ELSE 'OK' 	END AS OVERDUE, ZTBPOTable.REMARK, ZTBPOTable.PO_ID 	FROM 	(SELECT ZTBPOTable.EMPL_NO, ZTBPOTable.CUST_CD, ZTBPOTable.G_CODE, ZTBPOTable.PO_NO, isnull(SUM(ZTBDelivery.DELIVERY_QTY),0) AS TotalDelivered 	FROM ZTBPOTable  LEFT JOIN ZTBDelivery 	ON (ZTBDelivery.CTR_CD = ZTBPOTable.CTR_CD AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) 	GROUP BY ZTBPOTable.CTR_CD,ZTBPOTable.EMPL_NO,ZTBPOTable.G_CODE,ZTBPOTable.CUST_CD,ZTBPOTable.PO_NO) AS AA 	LEFT JOIN M010 	ON (M010.EMPL_NO = AA.EMPL_NO) 	LEFT JOIN M100 	ON (M100.G_CODE = AA.G_CODE) 	LEFT JOIN ZTBPOTable 	ON (AA.CUST_CD = ZTBPOTable.CUST_CD AND AA.G_CODE = ZTBPOTable.G_CODE AND AA.PO_NO = ZTBPOTable.PO_NO) 	JOIN M110 	ON (M110.CUST_CD = AA.CUST_CD) 	WHERE PO_DATE BETWEEN @startdate AND @enddate ) AS XX ) AS YY PIVOT (SUM(YY.PO_QTY) FOR YY.CUST_NAME_KD IN ([SEV],[SEVT],[SAMSUNG-ASIA])) AS PV ) AS KK JOIN ( SELECT @startdate AS START_DATE, @enddate AS END_DATE, SUM(MM.PO_QTY) AS TOTAL_PO_QTY FROM (SELECT AA.PO_NO,  M100.PROD_TYPE, M100.PROD_MAIN_MATERIAL, ZTBPOTable.PO_DATE,ZTBPOTable.RD_DATE, M110.CUST_NAME_KD, M100.G_NAME, M010.EMPL_NAME, AA.G_CODE, ZTBPOTable.PO_QTY, ZTBPOTable.PROD_PRICE, AA.TotalDelivered as TOTAL_DELIVERED, (ZTBPOTable.PO_QTY-AA.TotalDelivered) As PO_BALANCE,(ZTBPOTable.PO_QTY*ZTBPOTable.PROD_PRICE) As PO_AMOUNT , (AA.TotalDelivered*ZTBPOTable.PROD_PRICE) As DELIVERED_AMOUNT, ((ZTBPOTable.PO_QTY-AA.TotalDelivered)*ZTBPOTable.PROD_PRICE) As BALANCE_AMOUNT,DATEPART( MONTH, PO_DATE) AS POMONTH, DATEPART( ISOWK, PO_DATE) AS POWEEKNUM, YEAR(PO_DATE) As PO_YEAR, 	CASE 		WHEN (ZTBPOTable.RD_DATE < GETDATE()-1) AND ((ZTBPOTable.PO_QTY-AA.TotalDelivered) <>0) THEN 'OVER'    		ELSE 'OK' 	END AS OVERDUE, ZTBPOTable.REMARK, ZTBPOTable.PO_ID 	FROM 	(SELECT ZTBPOTable.EMPL_NO, ZTBPOTable.CUST_CD, ZTBPOTable.G_CODE, ZTBPOTable.PO_NO, isnull(SUM(ZTBDelivery.DELIVERY_QTY),0) AS TotalDelivered 	FROM ZTBPOTable  LEFT JOIN ZTBDelivery 	ON (ZTBDelivery.CTR_CD = ZTBPOTable.CTR_CD AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) 	GROUP BY ZTBPOTable.CTR_CD,ZTBPOTable.EMPL_NO,ZTBPOTable.G_CODE,ZTBPOTable.CUST_CD,ZTBPOTable.PO_NO) AS AA 	LEFT JOIN M010 	ON (M010.EMPL_NO = AA.EMPL_NO) 	LEFT JOIN M100 	ON (M100.G_CODE = AA.G_CODE) 	LEFT JOIN ZTBPOTable 	ON (AA.CUST_CD = ZTBPOTable.CUST_CD AND AA.G_CODE = ZTBPOTable.G_CODE AND AA.PO_NO = ZTBPOTable.PO_NO) 	JOIN M110 	ON (M110.CUST_CD = AA.CUST_CD) 	WHERE PO_DATE BETWEEN @startdate AND @enddate ) AS MM ) AS NN ON (KK.START_DATE = NN.START_DATE AND KK.END_DATE = NN.END_DATE)";
            result = config.GetData(strQuery);
            return result;

        }
    


        public DataTable report_WeeklyPOByCustomer()
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT  AA.PO_YEAR, AA.POWEEKNUM, BB.TOTAL_PO_QTY, AA.SEV, AA.SEVT, AA.SAMSUNG_ASIA,  (BB.TOTAL_PO_QTY-AA.SEV-AA.SEVT-AA.SAMSUNG_ASIA) AS OTHERS FROM ( 	SELECT  POWEEKNUM,  PO_YEAR, isnull(PV.[SEV],0) As SEV, isnull(PV.[SEVT],0) As SEVT,isnull(PV.[SAMSUNG-ASIA],0) As SAMSUNG_ASIA 	FROM 	( 	SELECT P.PO_QTY, P.CUST_NAME_KD,  P.POWEEKNUM, P.PO_YEAR FROM 	(   	SELECT AA.PO_NO,  M100.PROD_TYPE, M100.PROD_MAIN_MATERIAL, ZTBPOTable.PO_DATE,ZTBPOTable.RD_DATE, M110.CUST_NAME_KD, M100.G_NAME, M010.EMPL_NAME, AA.G_CODE, ZTBPOTable.PO_QTY, ZTBPOTable.PROD_PRICE, AA.TotalDelivered as TOTAL_DELIVERED, (ZTBPOTable.PO_QTY-AA.TotalDelivered) As PO_BALANCE,(ZTBPOTable.PO_QTY*ZTBPOTable.PROD_PRICE) As PO_AMOUNT , (AA.TotalDelivered*ZTBPOTable.PROD_PRICE) As DELIVERED_AMOUNT, ((ZTBPOTable.PO_QTY-AA.TotalDelivered)*ZTBPOTable.PROD_PRICE) As BALANCE_AMOUNT,DATEPART( MONTH, PO_DATE) AS POMONTH, DATEPART( ISOWK, PO_DATE) AS POWEEKNUM, YEAR(PO_DATE) As PO_YEAR, 	CASE 		WHEN (ZTBPOTable.RD_DATE < GETDATE()-1) AND ((ZTBPOTable.PO_QTY-AA.TotalDelivered) <>0) THEN 'OVER'    		ELSE 'OK' 	END AS OVERDUE, ZTBPOTable.REMARK, ZTBPOTable.PO_ID 	FROM 	(SELECT ZTBPOTable.EMPL_NO, ZTBPOTable.CUST_CD, ZTBPOTable.G_CODE, ZTBPOTable.PO_NO, isnull(SUM(ZTBDelivery.DELIVERY_QTY),0) AS TotalDelivered 	FROM ZTBPOTable  LEFT JOIN ZTBDelivery 	ON (ZTBDelivery.CTR_CD = ZTBPOTable.CTR_CD AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) 	GROUP BY ZTBPOTable.CTR_CD,ZTBPOTable.EMPL_NO,ZTBPOTable.G_CODE,ZTBPOTable.CUST_CD,ZTBPOTable.PO_NO) AS AA 	LEFT JOIN M010 	ON (M010.EMPL_NO = AA.EMPL_NO) 	LEFT JOIN M100 	ON (M100.G_CODE = AA.G_CODE) 	LEFT JOIN ZTBPOTable 	ON (AA.CUST_CD = ZTBPOTable.CUST_CD AND AA.G_CODE = ZTBPOTable.G_CODE AND AA.PO_NO = ZTBPOTable.PO_NO) 	JOIN M110 	ON (M110.CUST_CD = AA.CUST_CD)    	) AS P 	) AS j 	PIVOT (SUM(j.PO_QTY) FOR j.CUST_NAME_KD IN ([SEV],[SEVT],[SAMSUNG-ASIA])) AS PV ) AS AA JOIN (SELECT A.POWEEKNUM, SUM(A.PO_QTY) As TOTAL_PO_QTY FROM (	SELECT AA.PO_NO,  M100.PROD_TYPE, M100.PROD_MAIN_MATERIAL, ZTBPOTable.PO_DATE,ZTBPOTable.RD_DATE, M110.CUST_NAME_KD, M100.G_NAME, M010.EMPL_NAME, AA.G_CODE, ZTBPOTable.PO_QTY, ZTBPOTable.PROD_PRICE, AA.TotalDelivered as TOTAL_DELIVERED, (ZTBPOTable.PO_QTY-AA.TotalDelivered) As PO_BALANCE,(ZTBPOTable.PO_QTY*ZTBPOTable.PROD_PRICE) As PO_AMOUNT , (AA.TotalDelivered*ZTBPOTable.PROD_PRICE) As DELIVERED_AMOUNT, ((ZTBPOTable.PO_QTY-AA.TotalDelivered)*ZTBPOTable.PROD_PRICE) As BALANCE_AMOUNT,DATEPART( MONTH, PO_DATE) AS POMONTH, DATEPART( ISOWK, PO_DATE) AS POWEEKNUM, YEAR(PO_DATE) As PO_YEAR, 	CASE 		WHEN (ZTBPOTable.RD_DATE < GETDATE()-1) AND ((ZTBPOTable.PO_QTY-AA.TotalDelivered) <>0) THEN 'OVER'    		ELSE 'OK' 	END AS OVERDUE, ZTBPOTable.REMARK, ZTBPOTable.PO_ID 	FROM 	(SELECT ZTBPOTable.EMPL_NO, ZTBPOTable.CUST_CD, ZTBPOTable.G_CODE, ZTBPOTable.PO_NO, isnull(SUM(ZTBDelivery.DELIVERY_QTY),0) AS TotalDelivered 	FROM ZTBPOTable  LEFT JOIN ZTBDelivery 	ON (ZTBDelivery.CTR_CD = ZTBPOTable.CTR_CD AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) 	GROUP BY ZTBPOTable.CTR_CD,ZTBPOTable.EMPL_NO,ZTBPOTable.G_CODE,ZTBPOTable.CUST_CD,ZTBPOTable.PO_NO) AS AA 	LEFT JOIN M010 	ON (M010.EMPL_NO = AA.EMPL_NO) 	LEFT JOIN M100 	ON (M100.G_CODE = AA.G_CODE) 	LEFT JOIN ZTBPOTable 	ON (AA.CUST_CD = ZTBPOTable.CUST_CD AND AA.G_CODE = ZTBPOTable.G_CODE AND AA.PO_NO = ZTBPOTable.PO_NO) 	JOIN M110 	ON (M110.CUST_CD = AA.CUST_CD) ) AS A GROUP BY POWEEKNUM ) AS BB ON AA.POWEEKNUM = BB.POWEEKNUM ORDER BY PO_YEAR DESC,POWEEKNUM DESC";
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable updateCustomer(string CUST_CD, string CUST_NAME_KD)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "UPDATE M110 SET CUST_NAME_KD='" + CUST_NAME_KD +  "'  WHERE  CUST_CD='" + CUST_CD + "'";
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable updateOnline(string EMPL_NO)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = $"UPDATE M010 SET ONLINE_DATETIME=GETDATE() WHERE EMPL_NO='{EMPL_NO}'";
            result = config.GetData(strQuery);
            return result;

        }


        public DataTable updateInfo(string G_CODE, string G_CODE_KD, string PROD_TYPE,string PROD_MODEL, string PROD_PROJECT, string PROD_MAIN_MATERIAL)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "UPDATE M100 SET PROD_TYPE='" + PROD_TYPE + " ', PROD_MODEL='" + PROD_MODEL +" ', PROD_PROJECT='" +  PROD_PROJECT + "', PROD_MAIN_MATERIAL='" + PROD_MAIN_MATERIAL + "', G_NAME_KD='" + G_CODE_KD + "' WHERE  G_CODE='" + G_CODE + "'";
            result = config.GetData(strQuery);
            return result;

        }
        public DataTable updateInfoQLSX(string G_CODE, string FACTORY, string EQ1, string EQ2, string SETTING1, string SETTING2, string UPH1, string UPH2, string STEP1, string STEP2, string NOTE)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "UPDATE M100 SET FACTORY = '" + FACTORY + "', EQ1 = '" + EQ1 + "', EQ2 = '" + EQ2 + "', SETTING1 = '" + SETTING1 + "', SETTING2 = '" + SETTING2 + "', UPH1 = '" + UPH1 + "', UPH2 = '" + UPH2 + "', STEP1 = '" + STEP1 + "',STEP2 = '" + STEP2 + "', NOTE = '" + NOTE + "' WHERE G_CODE = '" + G_CODE + "'";
            //MessageBox.Show(strQuery);
            result = config.GetData(strQuery);
            return result;

        }


        public DataTable traPOTotal(string condition)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            //string strQuery = "SELECT  SUM(ZTBPOTable.PO_QTY) As PO_QTY, SUM(AA.TotalDelivered) as TOTAL_DELIVERED, SUM((ZTBPOTable.PO_QTY-AA.TotalDelivered)) As PO_BALANCE,SUM((ZTBPOTable.PO_QTY*ZTBPOTable.PROD_PRICE)) As PO_AMOUNT , SUM((AA.TotalDelivered*ZTBPOTable.PROD_PRICE)) As DELIVERED_AMOUNT, SUM(((ZTBPOTable.PO_QTY-AA.TotalDelivered)*ZTBPOTable.PROD_PRICE)) As BALANCE_AMOUNT FROM (SELECT ZTBPOTable.EMPL_NO, ZTBPOTable.CUST_CD, ZTBPOTable.G_CODE, ZTBPOTable.PO_NO, isnull(SUM(ZTBDelivery.DELIVERY_QTY),0) AS TotalDelivered FROM ZTBPOTable  LEFT JOIN ZTBDelivery  ON (ZTBDelivery.CTR_CD = ZTBPOTable.CTR_CD AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO)  GROUP BY ZTBPOTable.CTR_CD,ZTBPOTable.EMPL_NO,ZTBPOTable.G_CODE,ZTBPOTable.CUST_CD,ZTBPOTable.PO_NO) AS AA LEFT JOIN M010 ON (M010.EMPL_NO = AA.EMPL_NO) LEFT JOIN M100 ON (M100.G_CODE = AA.G_CODE) LEFT JOIN ZTBPOTable ON (AA.CUST_CD = ZTBPOTable.CUST_CD AND AA.G_CODE = ZTBPOTable.G_CODE AND AA.PO_NO = ZTBPOTable.PO_NO)  JOIN M110 ON (M110.CUST_CD = AA.CUST_CD) " + condition;
            string strQuery = "SELECT  SUM(cast(ZTBPOTable.PO_QTY as bigint)) As PO_QTY, SUM(cast(AA.TotalDelivered as bigint)) as TOTAL_DELIVERED, SUM(cast((ZTBPOTable.PO_QTY-AA.TotalDelivered) as bigint)) As PO_BALANCE,SUM((ZTBPOTable.PO_QTY*ZTBPOTable.PROD_PRICE)) As PO_AMOUNT , SUM((AA.TotalDelivered*ZTBPOTable.PROD_PRICE)) As DELIVERED_AMOUNT, SUM(((ZTBPOTable.PO_QTY-AA.TotalDelivered)*ZTBPOTable.PROD_PRICE)) As BALANCE_AMOUNT FROM (SELECT ZTBPOTable.EMPL_NO, ZTBPOTable.CUST_CD, ZTBPOTable.G_CODE, ZTBPOTable.PO_NO, isnull(SUM(ZTBDelivery.DELIVERY_QTY),0) AS TotalDelivered FROM ZTBPOTable  LEFT JOIN ZTBDelivery  ON (ZTBDelivery.CTR_CD = ZTBPOTable.CTR_CD AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO)  GROUP BY ZTBPOTable.CTR_CD,ZTBPOTable.EMPL_NO,ZTBPOTable.G_CODE,ZTBPOTable.CUST_CD,ZTBPOTable.PO_NO) AS AA LEFT JOIN M010 ON (M010.EMPL_NO = AA.EMPL_NO) LEFT JOIN M100 ON (M100.G_CODE = AA.G_CODE) LEFT JOIN ZTBPOTable ON (AA.CUST_CD = ZTBPOTable.CUST_CD AND AA.G_CODE = ZTBPOTable.G_CODE AND AA.PO_NO = ZTBPOTable.PO_NO)  JOIN M110 ON (M110.CUST_CD = AA.CUST_CD) " + condition;


            result = config.GetData(strQuery);
            return result;

        }

        public int checkPOBalance(string CUST_CD, string G_CODE, string PO_NO)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT  (ZTBPOTable.PO_QTY-AA.TotalDelivered) As PO_BALANCE FROM (SELECT ZTBPOTable.EMPL_NO, ZTBPOTable.CUST_CD, ZTBPOTable.G_CODE, ZTBPOTable.PO_NO, isnull(SUM(ZTBDelivery.DELIVERY_QTY),0) AS TotalDelivered FROM ZTBPOTable  LEFT JOIN ZTBDelivery ON (ZTBDelivery.CTR_CD = ZTBPOTable.CTR_CD AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) GROUP BY ZTBPOTable.CTR_CD,ZTBPOTable.EMPL_NO,ZTBPOTable.G_CODE,ZTBPOTable.CUST_CD,ZTBPOTable.PO_NO) AS AA LEFT JOIN M010 ON (M010.EMPL_NO = AA.EMPL_NO) LEFT JOIN M100 ON (M100.G_CODE = AA.G_CODE) LEFT JOIN ZTBPOTable ON (AA.CUST_CD = ZTBPOTable.CUST_CD AND AA.G_CODE = ZTBPOTable.G_CODE AND AA.PO_NO = ZTBPOTable.PO_NO) JOIN M110 ON (M110.CUST_CD = AA.CUST_CD) WHERE ZTBPOTable.G_CODE='" + G_CODE + "' AND ZTBPOTable.CUST_CD='" + CUST_CD + "' AND ZTBPOTable.PO_NO='"+ PO_NO +"'";
            result = config.GetData(strQuery);
            int po_balance = 0;
            if(result.Rows.Count >0)
            {
                po_balance = int.Parse(result.Rows[0]["PO_BALANCE"].ToString());
            }
            else
            {
                po_balance = 0;
            }
            
            return po_balance;
        }

        public int checkPOExist(string CUST_CD, string G_CODE, string PO_NO)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT  PO_ID FROM ZTBPOTable WHERE ZTBPOTable.G_CODE='" + G_CODE + "' AND ZTBPOTable.CUST_CD='" + CUST_CD + "' AND ZTBPOTable.PO_NO='" + PO_NO + "'";
            result = config.GetData(strQuery);
            int po_id;
            try
            {
                po_id = int.Parse(result.Rows[0]["PO_ID"].ToString());
            }
            catch(Exception ex)
            {
                po_id = -1;
            }
            //MessageBox.Show("PO_ID" +   po_id);
            return po_id;
        }

        

        public int checkFCSTExist(string CUST_CD, string G_CODE, string FCSTYEAR, string FCSTWEEKNO)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT  FCST_ID FROM ZTBFCSTTB WHERE ZTBFCSTTB.G_CODE='" + G_CODE + "' AND ZTBFCSTTB.CUST_CD='" + CUST_CD + "' AND ZTBFCSTTB.FCSTYEAR='" + FCSTYEAR + "' AND ZTBFCSTTB.FCSTWEEKNO='" + FCSTWEEKNO  + "'";
            result = config.GetData(strQuery);
            int po_id;
            try
            {
                po_id = int.Parse(result.Rows[0]["FCST_ID"].ToString());
            }
            catch (Exception ex)
            {
                po_id = -1;
            }
            return po_id;
        }

        public int checkKHGHExist(string CUST_CD, string G_CODE, string PLAN_DATE)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT  PLAN_ID FROM ZTBPLANTB WHERE ZTBPLANTB.G_CODE='" + G_CODE + "' AND ZTBPLANTB.CUST_CD='" + CUST_CD + "' AND ZTBPLANTB.PLAN_DATE='" + PLAN_DATE + "'";
            result = config.GetData(strQuery);
            int plan_id;
            try
            {
                plan_id = int.Parse(result.Rows[0]["PLAN_ID"].ToString());
            }
            catch (Exception ex)
            {
                plan_id = -1;
            }
            return plan_id;
        }



        public DataTable traPO(string condition)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();

            //string strQuery = $"SELECT  ZTBPOTable.PO_ID, M110.CUST_NAME_KD,AA.PO_NO,   M100.G_NAME,M100.G_NAME_KD, AA.G_CODE, ZTBPOTable.PO_DATE, ZTBPOTable.RD_DATE, ZTBPOTable.PROD_PRICE, ZTBPOTable.PO_QTY, AA.TotalDelivered as TOTAL_DELIVERED, (ZTBPOTable.PO_QTY-AA.TotalDelivered) As PO_BALANCE,(ZTBPOTable.PO_QTY*ZTBPOTable.PROD_PRICE) As PO_AMOUNT , (AA.TotalDelivered*ZTBPOTable.PROD_PRICE) As DELIVERED_AMOUNT, ((ZTBPOTable.PO_QTY-AA.TotalDelivered)*ZTBPOTable.PROD_PRICE) As BALANCE_AMOUNT, M010.EMPL_NAME, M100.PROD_TYPE, KKK.M_NAME_FULLBOM,M100.PROD_MAIN_MATERIAL, M110.CUST_CD,  M010.EMPL_NO,    DATEPART( MONTH, PO_DATE) AS POMONTH, DATEPART( ISOWK, PO_DATE) AS POWEEKNUM, CASE WHEN (ZTBPOTable.RD_DATE < GETDATE()-1) AND ((ZTBPOTable.PO_QTY-AA.TotalDelivered) <>0) THEN 'OVER'  ELSE 'OK' END AS OVERDUE, ZTBPOTable.REMARK FROM (SELECT ZTBPOTable.EMPL_NO, ZTBPOTable.CUST_CD, ZTBPOTable.G_CODE, ZTBPOTable.PO_NO, isnull(SUM(ZTBDelivery.DELIVERY_QTY),0) AS TotalDelivered FROM ZTBPOTable LEFT JOIN ZTBDelivery ON (ZTBDelivery.CTR_CD = ZTBPOTable.CTR_CD AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) GROUP BY ZTBPOTable.CTR_CD,ZTBPOTable.EMPL_NO,ZTBPOTable.G_CODE,ZTBPOTable.CUST_CD,ZTBPOTable.PO_NO) AS AA LEFT JOIN M010 ON (M010.EMPL_NO = AA.EMPL_NO) LEFT JOIN M100 ON (M100.G_CODE = AA.G_CODE) LEFT JOIN ZTBPOTable ON (AA.CUST_CD = ZTBPOTable.CUST_CD AND AA.G_CODE = ZTBPOTable.G_CODE AND AA.PO_NO = ZTBPOTable.PO_NO) LEFT JOIN M110 ON (M110.CUST_CD = AA.CUST_CD) LEFT JOIN (SELECT BBB.G_CODE, string_agg(BBB.M_NAME, ', ') AS M_NAME_FULLBOM FROM  (SELECT DISTINCT AAA.G_CODE, M090.M_NAME FROM 	( 		(SELECT DISTINCT G_CODE, M_CODE FROM M140) AS AAA LEFT JOIN M090 ON (AAA.M_CODE = M090.M_CODE) 	) ) AS BBB GROUP BY BBB.G_CODE ) AS KKK ON (KKK.G_CODE = ZTBPOTable.G_CODE) {condition} ";
            string strQuery = $"SELECT ZTBPOTable.PO_ID, M110.CUST_NAME_KD,AA.PO_NO, M100.G_NAME,M100.G_NAME_KD, AA.G_CODE, ZTBPOTable.PO_DATE, ZTBPOTable.RD_DATE, ZTBPOTable.PROD_PRICE, ZTBPOTable.PO_QTY, AA.TotalDelivered as TOTAL_DELIVERED, (ZTBPOTable.PO_QTY-AA.TotalDelivered) As PO_BALANCE,(ZTBPOTable.PO_QTY*ZTBPOTable.PROD_PRICE) As PO_AMOUNT , (AA.TotalDelivered*ZTBPOTable.PROD_PRICE) As DELIVERED_AMOUNT, ((ZTBPOTable.PO_QTY-AA.TotalDelivered)*ZTBPOTable.PROD_PRICE) As BALANCE_AMOUNT, isnull(TONKHOFULL.TONG_TON_KIEM,0) AS TON_KIEM , isnull(TONKHOFULL.BTP,0) AS BTP ,isnull(TONKHOFULL.TON_TP,0) AS TP , isnull(TONKHOFULL.BLOCK_QTY,0) AS BLOCK_QTY , isnull(TONKHOFULL.GRAND_TOTAL_STOCK,0) AS GRAND_TOTAL_STOCK , M010.EMPL_NAME, M100.PROD_TYPE, KKK.M_NAME_FULLBOM,M100.PROD_MAIN_MATERIAL, M110.CUST_CD, M010.EMPL_NO, DATEPART( MONTH, PO_DATE) AS POMONTH, DATEPART( ISOWK, PO_DATE) AS POWEEKNUM, CASE WHEN (ZTBPOTable.RD_DATE < GETDATE()-1) AND ((ZTBPOTable.PO_QTY-AA.TotalDelivered) <>0) THEN 'OVER' ELSE 'OK' END AS OVERDUE, ZTBPOTable.REMARK FROM (SELECT ZTBPOTable.EMPL_NO, ZTBPOTable.CUST_CD, ZTBPOTable.G_CODE, ZTBPOTable.PO_NO, isnull(SUM(ZTBDelivery.DELIVERY_QTY),0) AS TotalDelivered FROM ZTBPOTable LEFT JOIN ZTBDelivery ON (ZTBDelivery.CTR_CD = ZTBPOTable.CTR_CD AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) GROUP BY ZTBPOTable.CTR_CD,ZTBPOTable.EMPL_NO,ZTBPOTable.G_CODE,ZTBPOTable.CUST_CD,ZTBPOTable.PO_NO) AS AA LEFT JOIN M010 ON (M010.EMPL_NO = AA.EMPL_NO) LEFT JOIN M100 ON (M100.G_CODE = AA.G_CODE) LEFT JOIN ZTBPOTable ON (AA.CUST_CD = ZTBPOTable.CUST_CD AND AA.G_CODE = ZTBPOTable.G_CODE AND AA.PO_NO = ZTBPOTable.PO_NO) LEFT JOIN M110 ON (M110.CUST_CD = AA.CUST_CD) LEFT JOIN (SELECT BBB.G_CODE, string_agg(BBB.M_NAME, ', ') AS M_NAME_FULLBOM FROM (SELECT DISTINCT AAA.G_CODE, M090.M_NAME FROM ( (SELECT DISTINCT G_CODE, M_CODE FROM M140) AS AAA LEFT JOIN M090 ON (AAA.M_CODE = M090.M_CODE) ) ) AS BBB GROUP BY BBB.G_CODE ) AS KKK ON (KKK.G_CODE = ZTBPOTable.G_CODE) LEFT JOIN ( SELECT M100.G_CODE, M100.G_NAME, M100.G_NAME_KD, isnull(TONKIEM.INSPECT_BALANCE_QTY,0) AS CHO_KIEM, isnull(TONKIEM.WAIT_CS_QTY,0) AS CHO_CS_CHECK,isnull(TONKIEM.WAIT_SORTING_RMA,0) CHO_KIEM_RMA, isnull(TONKIEM.TOTAL_WAIT,0) AS TONG_TON_KIEM, isnull(BTP.BTP_QTY_EA,0) AS BTP, isnull(THANHPHAM.TONKHO,0) AS TON_TP, isnull(tbl_Block_table2.Block_Qty,0) AS BLOCK_QTY, (isnull(TONKIEM.TOTAL_WAIT,0) + isnull(BTP.BTP_QTY_EA,0)+ isnull(THANHPHAM.TONKHO,0) - isnull(tbl_Block_table2.Block_Qty,0)) AS GRAND_TOTAL_STOCK FROM M100 LEFT JOIN ( SELECT Product_MaVach, isnull([IN],0) AS NHAPKHO, isnull([OUT],0) AS XUATKHO, (isnull([IN],0)- isnull([OUT],0)) AS TONKHO FROM ( SELECT Product_Mavach, IO_Type, IO_Qty FROM tbl_InputOutput ) AS SourceTable PIVOT ( SUM(IO_Qty) FOR IO_Type IN ([IN], [OUT]) ) AS PivotTable ) AS THANHPHAM ON (THANHPHAM.Product_MaVach = M100.G_CODE) LEFT JOIN ( SELECT ZTB_WAIT_INSPECT.G_CODE, M100.G_NAME, M100.G_NAME_KD, SUM(INSPECT_BALANCE_QTY) AS INSPECT_BALANCE_QTY, SUM(WAIT_CS_QTY) AS WAIT_CS_QTY, SUM(WAIT_SORTING_RMA) AS WAIT_SORTING_RMA, SUM(INSPECT_BALANCE_QTY+ WAIT_CS_QTY+ WAIT_SORTING_RMA) AS TOTAL_WAIT FROM ZTB_WAIT_INSPECT JOIN M100 ON ( M100.G_CODE = ZTB_WAIT_INSPECT.G_CODE) WHERE UPDATE_DATE=CONVERT(date,GETDATE()) AND CALAMVIEC = 'DEM' GROUP BY ZTB_WAIT_INSPECT.G_CODE, M100.G_NAME, M100.G_NAME_KD) AS TONKIEM ON (THANHPHAM.Product_MaVach = TONKIEM.G_CODE) LEFT JOIN ( SELECT Product_MaVach, SUM(Block_Qty) AS Block_Qty from tbl_Block2 GROUP BY Product_MaVach ) AS tbl_Block_table2 ON (tbl_Block_table2.Product_MaVach= M100.G_CODE) LEFT JOIN ( SELECT ZTB_HALF_GOODS.G_CODE, M100.G_NAME, SUM(BTP_QTY_EA) AS BTP_QTY_EA FROM ZTB_HALF_GOODS JOIN M100 ON (M100.G_CODE = ZTB_HALF_GOODS.G_CODE) WHERE UPDATE_DATE = CONVERT(date,GETDATE()) GROUP BY ZTB_HALF_GOODS.G_CODE, M100.G_NAME) AS BTP ON (BTP.G_CODE = THANHPHAM.Product_MaVach) ) AS TONKHOFULL ON (TONKHOFULL.G_CODE = ZTBPOTable.G_CODE) {condition}";

            //MessageBox.Show(strQuery);
            result = config.GetData(strQuery);
            return result;
        }

        public DataTable traPO_TONKHO(string condition)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();

            string strQuery = $" SELECT PO_TABLE_1.G_CODE,TONKHOFULL.G_NAME,TONKHOFULL.G_NAME_KD, PO_TABLE_1.PO_QTY, TOTAL_DELIVERED, PO_TABLE_1.PO_BALANCE, TONKHOFULL.CHO_KIEM, TONKHOFULL.CHO_CS_CHECK, TONKHOFULL.CHO_KIEM_RMA, TONKHOFULL.TONG_TON_KIEM, TONKHOFULL.BTP, TONKHOFULL.TON_TP, TONKHOFULL.BLOCK_QTY, TONKHOFULL.GRAND_TOTAL_STOCK, (TONKHOFULL.GRAND_TOTAL_STOCK-PO_TABLE_1.PO_BALANCE) AS THUA_THIEU FROM ( SELECT G_CODE, SUM(PO_QTY) AS PO_QTY, SUM(TOTAL_DELIVERED) AS TOTAL_DELIVERED, SUM(PO_BALANCE) AS PO_BALANCE FROM ( SELECT AA.G_CODE, ZTBPOTable.PO_QTY, AA.TotalDelivered as TOTAL_DELIVERED, (ZTBPOTable.PO_QTY-AA.TotalDelivered) As PO_BALANCE FROM (SELECT ZTBPOTable.EMPL_NO, ZTBPOTable.CUST_CD, ZTBPOTable.G_CODE, ZTBPOTable.PO_NO, isnull(SUM(ZTBDelivery.DELIVERY_QTY),0) AS TotalDelivered FROM ZTBPOTable LEFT JOIN ZTBDelivery ON (ZTBDelivery.CTR_CD = ZTBPOTable.CTR_CD AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) GROUP BY ZTBPOTable.CTR_CD,ZTBPOTable.EMPL_NO,ZTBPOTable.G_CODE,ZTBPOTable.CUST_CD,ZTBPOTable.PO_NO) AS AA LEFT JOIN ZTBPOTable ON (AA.CUST_CD = ZTBPOTable.CUST_CD AND AA.G_CODE = ZTBPOTable.G_CODE AND AA.PO_NO = ZTBPOTable.PO_NO) ) AS PO_BALANCE_TABLE GROUP BY G_CODE ) AS PO_TABLE_1 LEFT JOIN ( SELECT M100.G_CODE, M100.G_NAME, M100.G_NAME_KD, isnull(TONKIEM.INSPECT_BALANCE_QTY,0) AS CHO_KIEM, isnull(TONKIEM.WAIT_CS_QTY,0) AS CHO_CS_CHECK,isnull(TONKIEM.WAIT_SORTING_RMA,0) CHO_KIEM_RMA, isnull(TONKIEM.TOTAL_WAIT,0) AS TONG_TON_KIEM, isnull(BTP.BTP_QTY_EA,0) AS BTP, isnull(THANHPHAM.TONKHO,0) AS TON_TP, isnull(tbl_Block_table2.Block_Qty,0) AS BLOCK_QTY, (isnull(TONKIEM.TOTAL_WAIT,0) + isnull(BTP.BTP_QTY_EA,0)+ isnull(THANHPHAM.TONKHO,0) - isnull(tbl_Block_table2.Block_Qty,0)) AS GRAND_TOTAL_STOCK FROM M100 LEFT JOIN ( SELECT Product_MaVach, isnull([IN],0) AS NHAPKHO, isnull([OUT],0) AS XUATKHO, (isnull([IN],0)- isnull([OUT],0)) AS TONKHO FROM ( SELECT Product_Mavach, IO_Type, IO_Qty FROM tbl_InputOutput ) AS SourceTable PIVOT ( SUM(IO_Qty) FOR IO_Type IN ([IN], [OUT]) ) AS PivotTable ) AS THANHPHAM ON (THANHPHAM.Product_MaVach = M100.G_CODE) LEFT JOIN ( SELECT ZTB_WAIT_INSPECT.G_CODE, M100.G_NAME, M100.G_NAME_KD, SUM(INSPECT_BALANCE_QTY) AS INSPECT_BALANCE_QTY, SUM(WAIT_CS_QTY) AS WAIT_CS_QTY, SUM(WAIT_SORTING_RMA) AS WAIT_SORTING_RMA, SUM(INSPECT_BALANCE_QTY+ WAIT_CS_QTY+ WAIT_SORTING_RMA) AS TOTAL_WAIT FROM ZTB_WAIT_INSPECT JOIN M100 ON ( M100.G_CODE = ZTB_WAIT_INSPECT.G_CODE) WHERE UPDATE_DATE=CONVERT(date,GETDATE()) AND CALAMVIEC = 'DEM' GROUP BY ZTB_WAIT_INSPECT.G_CODE, M100.G_NAME, M100.G_NAME_KD) AS TONKIEM ON (THANHPHAM.Product_MaVach = TONKIEM.G_CODE) LEFT JOIN ( SELECT Product_MaVach, SUM(Block_Qty) AS Block_Qty from tbl_Block2 GROUP BY Product_MaVach ) AS tbl_Block_table2 ON (tbl_Block_table2.Product_MaVach= M100.G_CODE) LEFT JOIN ( SELECT ZTB_HALF_GOODS.G_CODE, M100.G_NAME, SUM(BTP_QTY_EA) AS BTP_QTY_EA FROM ZTB_HALF_GOODS JOIN M100 ON (M100.G_CODE = ZTB_HALF_GOODS.G_CODE) WHERE UPDATE_DATE = CONVERT(date,GETDATE()) GROUP BY ZTB_HALF_GOODS.G_CODE, M100.G_NAME) AS BTP ON (BTP.G_CODE = THANHPHAM.Product_MaVach) ) AS TONKHOFULL ON (TONKHOFULL.G_CODE = PO_TABLE_1.G_CODE) {condition}";


            //MessageBox.Show(strQuery);
            result = config.GetData(strQuery);
            return result;
        }

        public DataTable traPO_TONKHOKD(string condition)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();

            string strQuery = $"SELECT TONKHOFULL.G_NAME_KD, SUM(PO_TABLE_1.PO_QTY) AS PO_QTY , SUM(TOTAL_DELIVERED) AS TOTAL_DELIVERED, SUM(PO_TABLE_1.PO_BALANCE) AS PO_BALANCE, SUM(TONKHOFULL.CHO_KIEM) AS CHO_KIEM, SUM(TONKHOFULL.CHO_CS_CHECK) AS CHO_CS_CHECK, SUM(TONKHOFULL.CHO_KIEM_RMA) AS CHO_KIEM_RMA, SUM(TONKHOFULL.TONG_TON_KIEM) AS TONG_TON_KIEM, SUM(TONKHOFULL.BTP) AS BTP, SUM(TONKHOFULL.TON_TP) AS TON_TP, SUM(TONKHOFULL.BLOCK_QTY) AS BLOCK_QTY, SUM(TONKHOFULL.GRAND_TOTAL_STOCK) AS GRAND_TOTAL_STOCK, SUM((TONKHOFULL.GRAND_TOTAL_STOCK-PO_TABLE_1.PO_BALANCE)) AS THUA_THIEU FROM ( SELECT G_CODE, SUM(PO_QTY) AS PO_QTY, SUM(TOTAL_DELIVERED) AS TOTAL_DELIVERED, SUM(PO_BALANCE) AS PO_BALANCE FROM ( SELECT AA.G_CODE, ZTBPOTable.PO_QTY, AA.TotalDelivered as TOTAL_DELIVERED, (ZTBPOTable.PO_QTY-AA.TotalDelivered) As PO_BALANCE FROM (SELECT ZTBPOTable.EMPL_NO, ZTBPOTable.CUST_CD, ZTBPOTable.G_CODE, ZTBPOTable.PO_NO, isnull(SUM(ZTBDelivery.DELIVERY_QTY),0) AS TotalDelivered FROM ZTBPOTable LEFT JOIN ZTBDelivery ON (ZTBDelivery.CTR_CD = ZTBPOTable.CTR_CD AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) GROUP BY ZTBPOTable.CTR_CD,ZTBPOTable.EMPL_NO,ZTBPOTable.G_CODE,ZTBPOTable.CUST_CD,ZTBPOTable.PO_NO) AS AA LEFT JOIN ZTBPOTable ON (AA.CUST_CD = ZTBPOTable.CUST_CD AND AA.G_CODE = ZTBPOTable.G_CODE AND AA.PO_NO = ZTBPOTable.PO_NO) ) AS PO_BALANCE_TABLE GROUP BY G_CODE ) AS PO_TABLE_1 LEFT JOIN ( SELECT M100.G_CODE, M100.G_NAME, M100.G_NAME_KD, isnull(TONKIEM.INSPECT_BALANCE_QTY,0) AS CHO_KIEM, isnull(TONKIEM.WAIT_CS_QTY,0) AS CHO_CS_CHECK,isnull(TONKIEM.WAIT_SORTING_RMA,0) CHO_KIEM_RMA, isnull(TONKIEM.TOTAL_WAIT,0) AS TONG_TON_KIEM, isnull(BTP.BTP_QTY_EA,0) AS BTP, isnull(THANHPHAM.TONKHO,0) AS TON_TP, isnull(tbl_Block_table2.Block_Qty,0) AS BLOCK_QTY, (isnull(TONKIEM.TOTAL_WAIT,0) + isnull(BTP.BTP_QTY_EA,0)+ isnull(THANHPHAM.TONKHO,0) - isnull(tbl_Block_table2.Block_Qty,0)) AS GRAND_TOTAL_STOCK FROM M100 LEFT JOIN ( SELECT Product_MaVach, isnull([IN],0) AS NHAPKHO, isnull([OUT],0) AS XUATKHO, (isnull([IN],0)- isnull([OUT],0)) AS TONKHO FROM ( SELECT Product_Mavach, IO_Type, IO_Qty FROM tbl_InputOutput ) AS SourceTable PIVOT ( SUM(IO_Qty) FOR IO_Type IN ([IN], [OUT]) ) AS PivotTable ) AS THANHPHAM ON (THANHPHAM.Product_MaVach = M100.G_CODE) LEFT JOIN ( SELECT ZTB_WAIT_INSPECT.G_CODE, M100.G_NAME, M100.G_NAME_KD, SUM(INSPECT_BALANCE_QTY) AS INSPECT_BALANCE_QTY, SUM(WAIT_CS_QTY) AS WAIT_CS_QTY, SUM(WAIT_SORTING_RMA) AS WAIT_SORTING_RMA, SUM(INSPECT_BALANCE_QTY+ WAIT_CS_QTY+ WAIT_SORTING_RMA) AS TOTAL_WAIT FROM ZTB_WAIT_INSPECT JOIN M100 ON ( M100.G_CODE = ZTB_WAIT_INSPECT.G_CODE) WHERE UPDATE_DATE=CONVERT(date,GETDATE()) AND CALAMVIEC = 'DEM' GROUP BY ZTB_WAIT_INSPECT.G_CODE, M100.G_NAME, M100.G_NAME_KD) AS TONKIEM ON (THANHPHAM.Product_MaVach = TONKIEM.G_CODE) LEFT JOIN ( SELECT Product_MaVach, SUM(Block_Qty) AS Block_Qty from tbl_Block2 GROUP BY Product_MaVach ) AS tbl_Block_table2 ON (tbl_Block_table2.Product_MaVach= M100.G_CODE) LEFT JOIN ( SELECT ZTB_HALF_GOODS.G_CODE, M100.G_NAME, SUM(BTP_QTY_EA) AS BTP_QTY_EA FROM ZTB_HALF_GOODS JOIN M100 ON (M100.G_CODE = ZTB_HALF_GOODS.G_CODE) WHERE UPDATE_DATE = CONVERT(date,GETDATE()) GROUP BY ZTB_HALF_GOODS.G_CODE, M100.G_NAME) AS BTP ON (BTP.G_CODE = THANHPHAM.Product_MaVach) ) AS TONKHOFULL ON (TONKHOFULL.G_CODE = PO_TABLE_1.G_CODE)   {condition} GROUP BY TONKHOFULL.G_NAME_KD";

            //MessageBox.Show(strQuery);
            result = config.GetData(strQuery);
            return result;
        }

        public DataTable traFCST(string condition)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT ZTBFCSTTB.FCST_ID, ZTBFCSTTB.FCSTYEAR, ZTBFCSTTB.FCSTWEEKNO,ZTBFCSTTB.G_CODE, M100.G_NAME_KD, M100.G_NAME, M010.EMPL_NAME, M110.CUST_NAME_KD, M100.PROD_PROJECT, M100.PROD_MODEL, M100.PROD_MAIN_MATERIAL, ZTBFCSTTB.PROD_PRICE,ZTBFCSTTB.W1,ZTBFCSTTB.W2,ZTBFCSTTB.W3,ZTBFCSTTB.W4,ZTBFCSTTB.W5,ZTBFCSTTB.W6,ZTBFCSTTB.W7,ZTBFCSTTB.W8,ZTBFCSTTB.W9,ZTBFCSTTB.W10,ZTBFCSTTB.W11,ZTBFCSTTB.W12,ZTBFCSTTB.W13,ZTBFCSTTB.W14,ZTBFCSTTB.W15,ZTBFCSTTB.W16,ZTBFCSTTB.W17,ZTBFCSTTB.W18,ZTBFCSTTB.W19,ZTBFCSTTB.W20,ZTBFCSTTB.W21,ZTBFCSTTB.W22, ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W1 AS W1A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W2 AS W2A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W3  AS W3A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W4  AS W4A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W5 AS W5A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W6 AS W6A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W7 AS W7A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W8 AS W8A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W9 AS W9A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W10 AS W10A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W11 AS W11A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W12 AS W12A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W13 AS W13A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W14 AS W14A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W15 AS W15A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W16 AS W16A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W17 AS W17A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W18 AS W18A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W19 AS W19A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W20 AS W20A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W21 AS W21A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W22 AS W22A  FROM ZTBFCSTTB JOIN M100 ON (M100.G_CODE = ZTBFCSTTB.G_CODE) JOIN M110 ON(M110.CUST_CD = ZTBFCSTTB.CUST_CD) JOIN M010 ON (M010.EMPL_NO = ZTBFCSTTB.EMPL_NO)" + condition;
            result = config.GetData(strQuery);
            return result;
        }

        public DataTable traFCSTYERKWEEK(string fcst_year, string fcstweeknum)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT ZTBFCSTTB.FCST_ID, ZTBFCSTTB.FCSTYEAR, ZTBFCSTTB.FCSTWEEKNO,ZTBFCSTTB.G_CODE, M100.G_NAME_KD, M100.G_NAME, M010.EMPL_NAME, M110.CUST_NAME_KD, M100.PROD_PROJECT, M100.PROD_MODEL, M100.PROD_MAIN_MATERIAL, ZTBFCSTTB.PROD_PRICE,ZTBFCSTTB.W1,ZTBFCSTTB.W2,ZTBFCSTTB.W3,ZTBFCSTTB.W4,ZTBFCSTTB.W5,ZTBFCSTTB.W6,ZTBFCSTTB.W7,ZTBFCSTTB.W8,ZTBFCSTTB.W9,ZTBFCSTTB.W10,ZTBFCSTTB.W11,ZTBFCSTTB.W12,ZTBFCSTTB.W13,ZTBFCSTTB.W14,ZTBFCSTTB.W15,ZTBFCSTTB.W16,ZTBFCSTTB.W17,ZTBFCSTTB.W18,ZTBFCSTTB.W19,ZTBFCSTTB.W20,ZTBFCSTTB.W21,ZTBFCSTTB.W22, ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W1 AS W1A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W2 AS W2A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W3  AS W3A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W4  AS W4A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W5 AS W5A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W6 AS W6A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W7 AS W7A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W8 AS W8A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W9 AS W9A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W10 AS W10A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W11 AS W11A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W12 AS W12A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W13 AS W13A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W14 AS W14A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W15 AS W15A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W16 AS W16A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W17 AS W17A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W18 AS W18A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W19 AS W19A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W20 AS W20A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W21 AS W21A,ZTBFCSTTB.PROD_PRICE * ZTBFCSTTB.W22 AS W22A  FROM ZTBFCSTTB JOIN M100 ON (M100.G_CODE = ZTBFCSTTB.G_CODE) JOIN M110 ON(M110.CUST_CD = ZTBFCSTTB.CUST_CD) JOIN M010 ON (M010.EMPL_NO = ZTBFCSTTB.EMPL_NO)";
            result = config.GetData(strQuery);
            return result;
        }

        public DataTable traPlan(string condition)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT ZTBPLANTB.PLAN_ID, M010.EMPL_NAME, M110.CUST_NAME_KD, ZTBPLANTB.CUST_CD, ZTBPLANTB.G_CODE, M100.G_NAME_KD, M100.G_NAME,  M100.PROD_TYPE ,M100.PROD_MAIN_MATERIAL, ZTBPLANTB.PLAN_DATE, ZTBPLANTB.D1,ZTBPLANTB.D2,ZTBPLANTB.D3,ZTBPLANTB.D4,ZTBPLANTB.D5,ZTBPLANTB.D6,ZTBPLANTB.D7,ZTBPLANTB.D8, ZTBPLANTB.REMARK  FROM ZTBPLANTB JOIN M100 ON (M100.G_CODE = ZTBPLANTB.G_CODE) JOIN M110 ON (M110.CUST_CD = ZTBPLANTB.CUST_CD) JOIN M010 ON (M010.EMPL_NO= ZTBPLANTB.EMPL_NO)" + condition;
            result = config.GetData(strQuery);
            return result;
        }


        public DataTable customerDaily(string startdate, string enddate)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            //string strQuery = "DECLARE @startdate date DECLARE @enddate date SET @startdate = '"+ startdate + "' SET @enddate='" + enddate +"'   DECLARE @cols varchar(max) SELECT @cols= (SELECT DISTINCT  DELIVERY_DATE  FROM ZTBDelivery WHERE DELIVERY_DATE BETWEEN + @startdate  AND  @enddate ORDER BY DELIVERY_DATE ASC for xml  path('')) select @cols = replace(@cols,'<DELIVERY_DATE>','[') select @cols = replace(@cols,'</DELIVERY_DATE>', '],') select @cols = left(@cols,len(@cols) -1) DECLARE @cols2 varchar(max) SELECT @cols2= (SELECT DISTINCT  DELIVERY_DATE  FROM ZTBDelivery WHERE DELIVERY_DATE BETWEEN @startdate AND  @enddate ORDER BY DELIVERY_DATE ASC for xml  path('')) select @cols2 = replace(@cols2,'<DELIVERY_DATE>','isnull([') select @cols2 = replace(@cols2,'</DELIVERY_DATE>', '],0) AS D, ') select @cols2 = left(@cols2,len(@cols2) -1) DECLARE @cols3 varchar(max) SELECT @cols3= (SELECT DISTINCT  DELIVERY_DATE  FROM ZTBDelivery WHERE DELIVERY_DATE BETWEEN @startdate AND  @enddate ORDER BY DELIVERY_DATE ASC for xml  path('')) select @cols3 = replace(@cols3,'<DELIVERY_DATE>','isnull([') select @cols3 = replace(@cols3,'</DELIVERY_DATE>', '],0) +') select @cols3 = left(@cols3,len(@cols3) -1) declare @query varchar(max) select @query = 'select CUST_NAME_KD, ('+ @cols3 + ') AS DELIVERED_AMOUNT,  '+ @cols2 + ' from (  select XX.CUST_NAME_KD, XX.DELIVERED_AMOUNT ,XX.DELIVERY_DATE   from  (SELECT ZTBDelivery.G_CODE, M010.EMPL_NAME, M110.CUST_NAME_KD, ZTBDelivery.DELIVERY_DATE, M100.G_NAME, M100.PROD_MAIN_MATERIAL, ZTBDelivery.DELIVERY_QTY, ZTBPOTable.PROD_PRICE, ZTBDelivery.PO_NO, (ZTBPOTable.PROD_PRICE * ZTBDelivery.DELIVERY_QTY) As DELIVERED_AMOUNT, M100.PROD_TYPE, DATEPART( MONTH, ZTBDelivery.DELIVERY_DATE) AS DELMONTH, DATEPART( ISOWK,  ZTBDelivery.DELIVERY_DATE) AS DELWEEKNUM ,ZTBDelivery.NOCANCEL , ZTBDelivery.DELIVERY_ID  FROM ZTBDelivery JOIN ZTBPOTable ON (ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) JOIN M010 ON ZTBDelivery.EMPL_NO = M010.EMPL_NO JOIN M100 ON ZTBDelivery.G_CODE = M100.G_CODE JOIN M110 ON M110.CUST_CD = ZTBDelivery.CUST_CD  ) AS XX  ) src pivot (   SUM(DELIVERED_AMOUNT)   for DELIVERY_DATE in (' + @cols + ') ) piv;'  execute(@query)";
            
            string strQuery = "DECLARE @startdate date DECLARE @enddate date DECLARE @tempdate date  SET @startdate = '"+startdate + "' SET @enddate='" + enddate + "' SET @tempdate = @startdate   DECLARE @string varchar(max) DECLARE @string2 varchar(max) DECLARE @string3 varchar(max)  DECLARE @countdate int SET @countdate =0  SET @string ='' SET @string2='' SET @string3=''  WHILE @tempdate <= @enddate BEGIN 	SET @countdate = (SElECT COUNT(DELIVERY_DATE) AS tempdate FROM ZTBDelivery WHERE DELIVERY_DATE=@tempdate) 	IF (@countdate<>0) 	SELECT @string= @string +' isnull([' + CAST(@tempdate AS varchar(max))	+ '],0) AS [' + CAST(@tempdate AS varchar(max)) + '],' 	SET @tempdate =DATEADD(D,1, @tempdate) END SELECT @string = left(@string,len(@string) -1)   SET @tempdate = @startdate WHILE @tempdate <= @enddate BEGIN 	SELECT @string2= @string2 +'[' + CAST(@tempdate AS varchar(max))	+ '],'			 	SET @tempdate =DATEADD(D,1, @tempdate) 	 END SELECT @string2 = left(@string2,len(@string2) -1)   SET @tempdate = @startdate WHILE @tempdate <= @enddate BEGIN 	SELECT @string3= @string3 +'isnull([' + CAST(@tempdate AS varchar(max))	+ '],0) +' 	SET @tempdate =DATEADD(D,1, @tempdate) END SELECT @string3 = left(@string3,len(@string3) -1)    DECLARE @cols varchar(max) SELECT @cols= (SELECT DISTINCT  DELIVERY_DATE  FROM ZTBDelivery WHERE DELIVERY_DATE BETWEEN + @startdate  AND  @enddate ORDER BY DELIVERY_DATE ASC for xml  path('')) select @cols = replace(@cols,'<DELIVERY_DATE>','[') select @cols = replace(@cols,'</DELIVERY_DATE>', '],') select @cols = left(@cols,len(@cols) -1)   DECLARE @cols2 varchar(max) SELECT @cols2= (SELECT DISTINCT  DELIVERY_DATE  FROM ZTBDelivery WHERE DELIVERY_DATE BETWEEN @startdate AND  @enddate ORDER BY DELIVERY_DATE ASC for xml  path('')) select @cols2 = replace(@cols2,'<DELIVERY_DATE>','isnull([') select @cols2 = replace(@cols2,'</DELIVERY_DATE>', '],0) AS D, ') select @cols2 = left(@cols2,len(@cols2) -1) DECLARE @cols3 varchar(max) SELECT @cols3= (SELECT DISTINCT  DELIVERY_DATE  FROM ZTBDelivery WHERE DELIVERY_DATE BETWEEN @startdate AND  @enddate ORDER BY DELIVERY_DATE ASC for xml  path('')) select @cols3 = replace(@cols3,'<DELIVERY_DATE>','isnull([') select @cols3 = replace(@cols3,'</DELIVERY_DATE>', '],0) +') select @cols3 = left(@cols3,len(@cols3) -1) declare @query varchar(max) select @query = 'select CUST_NAME_KD, ('+ @cols3 + ') AS DELIVERED_AMOUNT,  '+@string + ' from (  select XX.CUST_NAME_KD, XX.DELIVERED_AMOUNT ,XX.DELIVERY_DATE   from  (SELECT ZTBDelivery.G_CODE, M010.EMPL_NAME, M110.CUST_NAME_KD, ZTBDelivery.DELIVERY_DATE, M100.G_NAME, M100.PROD_MAIN_MATERIAL, ZTBDelivery.DELIVERY_QTY, ZTBPOTable.PROD_PRICE, ZTBDelivery.PO_NO, (ZTBPOTable.PROD_PRICE * ZTBDelivery.DELIVERY_QTY) As DELIVERED_AMOUNT, M100.PROD_TYPE, DATEPART( MONTH, ZTBDelivery.DELIVERY_DATE) AS DELMONTH, DATEPART( ISOWK,  ZTBDelivery.DELIVERY_DATE) AS DELWEEKNUM ,ZTBDelivery.NOCANCEL , ZTBDelivery.DELIVERY_ID  FROM ZTBDelivery JOIN ZTBPOTable ON (ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) JOIN M010 ON ZTBDelivery.EMPL_NO = M010.EMPL_NO JOIN M100 ON ZTBDelivery.G_CODE = M100.G_CODE JOIN M110 ON M110.CUST_CD = ZTBDelivery.CUST_CD  ) AS XX  ) src pivot (   SUM(DELIVERED_AMOUNT)   for DELIVERY_DATE in (' + @cols + ') ) piv;'  execute(@query)";
            try
            {
                result = config.GetData(strQuery);
            }
            catch(Exception ex)
            {
               
            }

            return result;

        }


        public DataTable traInvoiceByPIC(string startdate, string enddate)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT M010.EMPL_NAME,  isnull(DELIVERYTB.TOTAL_DELIVERED,0) AS DELIVERY_QTY, isnull(TOTAL_DELIVERED_AMOUNT,0) AS DELIVERED_AMOUNT FROM (SELECT  M010.EMPL_NO,  SUM(ZTBDelivery.DELIVERY_QTY) AS TOTAL_DELIVERED, SUM((ZTBPOTable.PROD_PRICE * ZTBDelivery.DELIVERY_QTY)) As TOTAL_DELIVERED_AMOUNT FROM ZTBDelivery JOIN ZTBPOTable ON (ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) JOIN M010 ON ZTBDelivery.EMPL_NO = M010.EMPL_NO JOIN M100 ON ZTBDelivery.G_CODE = M100.G_CODE JOIN M110 ON M110.CUST_CD = ZTBDelivery.CUST_CD WHERE DELIVERY_DATE BETWEEN '" + startdate + "'  AND '" + enddate + "' GROUP By M010.EMPL_NO) AS DELIVERYTB JOIN M010 ON (M010.EMPL_NO = DELIVERYTB.EMPL_NO) ORDER BY TOTAL_DELIVERED_AMOUNT DESC";
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable traInvoice(string condition)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT ZTBDelivery.DELIVERY_ID,M110.CUST_NAME_KD, ZTBDelivery.PO_NO, ZTBDelivery.G_CODE,  M100.G_NAME, M100.G_NAME_KD,ZTBDelivery.DELIVERY_DATE, ZTBPOTable.PROD_PRICE,  ZTBDelivery.DELIVERY_QTY,  (ZTBPOTable.PROD_PRICE * ZTBDelivery.DELIVERY_QTY) As DELIVERED_AMOUNT, M010.EMPL_NAME,   M100.PROD_MAIN_MATERIAL,  M100.PROD_TYPE, DATEPART( MONTH, ZTBDelivery.DELIVERY_DATE) AS DELMONTH, DATEPART( ISOWK,  ZTBDelivery.DELIVERY_DATE) AS DELWEEKNUM ,ZTBDelivery.NOCANCEL ,ZTBDelivery.REMARK,  ZTBDelivery.INVOICE_NO, ZTBDelivery.CUST_CD, ZTBDelivery.EMPL_NO  FROM ZTBDelivery JOIN ZTBPOTable ON (ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) JOIN M010 ON ZTBDelivery.EMPL_NO = M010.EMPL_NO JOIN M100 ON ZTBDelivery.G_CODE = M100.G_CODE JOIN M110 ON M110.CUST_CD = ZTBDelivery.CUST_CD  " + condition;
            result = config.GetData(strQuery);
            return result;

        }
        public DataTable traInvoiceTotal(string condition)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
           // string strQuery = "SELECT SUM(ZTBDelivery.DELIVERY_QTY) AS DELIVERED_QTY, SUM((ZTBPOTable.PROD_PRICE * ZTBDelivery.DELIVERY_QTY)) As DELIVERED_AMOUNT, SUM(ZTBPOTable.PO_QTY) AS PO_QTY, (SUM(ZTBPOTable.PO_QTY) -  SUM(ZTBDelivery.DELIVERY_QTY)) AS PO_BALANCE FROM  ZTBDelivery JOIN ZTBPOTable ON (ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) JOIN M010 ON ZTBDelivery.EMPL_NO = M010.EMPL_NO JOIN M100 ON ZTBDelivery.G_CODE = M100.G_CODE JOIN M110 ON M110.CUST_CD = ZTBDelivery.CUST_CD " + condition;
            string strQuery = "SELECT SUM(cast(ZTBDelivery.DELIVERY_QTY as bigint)) AS DELIVERED_QTY, SUM((ZTBPOTable.PROD_PRICE * ZTBDelivery.DELIVERY_QTY)) As DELIVERED_AMOUNT, SUM(cast(ZTBPOTable.PO_QTY as bigint)) AS PO_QTY, (SUM(cast(ZTBPOTable.PO_QTY as  bigint)) -  SUM(cast(ZTBDelivery.DELIVERY_QTY as bigint))) AS PO_BALANCE FROM  ZTBDelivery JOIN ZTBPOTable ON (ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) JOIN M010 ON ZTBDelivery.EMPL_NO = M010.EMPL_NO JOIN M100 ON ZTBDelivery.G_CODE = M100.G_CODE JOIN M110 ON M110.CUST_CD = ZTBDelivery.CUST_CD " + condition;
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable updateINVOICE_NO(string CTR_CD, string CUST_CD, string EMPL_NO, string G_CODE, string PO_NO, string INVOICE_NO)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "UPDATE ZTBDelivery SET INVOICE_NO = '" + INVOICE_NO + "' WHERE CUST_CD='" + CUST_CD + "' AND EMPL_NO = '" + EMPL_NO + "' AND G_CODE='" + G_CODE + "' AND PO_NO = '" + PO_NO + "'";
            result = config.GetData(strQuery);
            return result;

        }



        public DataTable UpdateInvoice(string CTR_CD, string CUST_CD, string EMPL_NO, string G_CODE, string PO_NO, string DELIVERY_QTY, string DELIVERY_DATE, string NOCANCEL, string DELIVERY_ID)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "UPDATE ZTBDelivery  SET CTR_CD='002', CUST_CD='"+ CUST_CD + "', EMPL_NO = '"+ EMPL_NO+"', G_CODE='"+ G_CODE + "', PO_NO = '"+ PO_NO +"', DELIVERY_QTY= " + DELIVERY_QTY + " , DELIVERY_DATE='"+ DELIVERY_DATE + "' ,NOCANCEL = 1  WHERE DELIVERY_ID=" + DELIVERY_ID;
            result = config.GetData(strQuery);
            return result;
        }


        public DataTable InsertInvoice(string CTR_CD, string CUST_CD, string EMPL_NO, string G_CODE, string PO_NO, string DELIVERY_QTY, string DELIVERY_DATE, string NOCANCEL)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "INSERT INTO ZTBDelivery (CTR_CD, CUST_CD, EMPL_NO, G_CODE, PO_NO, DELIVERY_QTY, DELIVERY_DATE, NOCANCEL) VALUES ('" + CTR_CD + "','" + CUST_CD + "','" + EMPL_NO + "','" + G_CODE + "','" + PO_NO + "','" + DELIVERY_QTY + "','" + DELIVERY_DATE + "','" + NOCANCEL + "')";
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable traDelivery(string fromdate, string todate)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = $"SELECT AA.DELIVERY_DATE AS DELIVERY_DATE, AA.DELIVERY_QTY, BB.DELIVERY_AMOUNT FROM (SELECT DISTINCT DELIVERY_DATE, SUM(DELIVERY_QTY) AS DELIVERY_QTY FROM ZTBDelivery WHERE DELIVERY_DATE >= '{fromdate}' AND DELIVERY_DATE <= '{todate}' GROUP BY DELIVERY_DATE) AS AA JOIN (SELECT DISTINCT ZTBDelivery.DELIVERY_DATE, SUM(ZTBDelivery.DELIVERY_QTY * ZTBPOTable.PROD_PRICE) AS DELIVERY_AMOUNT FROM ZTBDelivery JOIN ZTBPOTable ON (ZTBDelivery.PO_NO = ZTBPOTable.PO_NO AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE)  WHERE ZTBDelivery.DELIVERY_DATE >= '{fromdate}' AND DELIVERY_DATE <= '{todate}' GROUP BY ZTBDelivery.DELIVERY_DATE) AS BB ON (AA.DELIVERY_DATE = BB.DELIVERY_DATE) ORDER BY AA.DELIVERY_DATE ASC";
            result = config.GetData(strQuery);
            return result;

        }
        public DataTable traDelivery_customer(string fromdate, string todate)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = $"SELECT TOP 5 * FROM (SELECT  DISTINCT M110.CUST_NAME_KD, SUM(ZTBDelivery.DELIVERY_QTY * ZTBPOTable.PROD_PRICE) AS DELIVERY_AMOUNT FROM ZTBDelivery JOIN ZTBPOTable ON (ZTBDelivery.PO_NO = ZTBPOTable.PO_NO AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE) LEFT JOIN M110 ON (M110.CUST_CD = ZTBDelivery.CUST_CD) WHERE ZTBDelivery.DELIVERY_DATE >= '{fromdate}' AND ZTBDelivery.DELIVERY_DATE <= '{todate}'GROUP BY M110.CUST_NAME_KD ) AS AA ORDER BY AA.DELIVERY_AMOUNT DESC";
            result = config.GetData(strQuery);
            return result;

        }
        public DataTable traDelivery_Amount()
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT DISTINCT ZTBDelivery.DELIVERY_DATE, SUM(ZTBDelivery.DELIVERY_QTY * ZTBPOTable.PROD_PRICE) AS DELIVERY_AMOUNT FROM ZTBDelivery JOIN ZTBPOTable ON (ZTBDelivery.PO_NO = ZTBPOTable.PO_NO AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE)  WHERE ZTBDelivery.DELIVERY_DATE >= '2022-06-01'GROUP BY ZTBDelivery.DELIVERY_DATE ORDER BY ZTBDelivery.DELIVERY_DATE ASC";
            result = config.GetData(strQuery);
            return result;

        }
        public int checkDeliveredQTy(string CUST_CD, string G_CODE, string PO_NO)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "SELECT (isnull(SUM(ZTBDelivery.DELIVERY_QTY),0)) AS TOTAL_DELIVERED 	FROM ZTBPOTable  LEFT JOIN ZTBDelivery 	ON (ZTBDelivery.CTR_CD = ZTBPOTable.CTR_CD AND ZTBDelivery.CUST_CD = ZTBPOTable.CUST_CD AND ZTBDelivery.G_CODE = ZTBPOTable.G_CODE AND ZTBDelivery.PO_NO = ZTBPOTable.PO_NO) 	WHERE ZTBPOTable.CUST_CD='" + CUST_CD + "' AND ZTBPOTable.G_CODE='" + G_CODE + "' AND ZTBPOTable.PO_NO='"+ PO_NO +"'";
            result = config.GetData(strQuery);
            int delivered_qty = 0;
            if (result.Rows.Count >0)
            {
                foreach (DataRow row in result.Rows)
                {
                    delivered_qty = int.Parse(row[0].ToString());
                }               
            }
            else
            {
                delivered_qty = 0;
            }            
            return delivered_qty;

        }

     


        public DataTable UpdatePO(string CTR_CD, string CUST_CD, string EMPL_NO, string G_CODE, string PO_NO, string PO_QTY, string PO_DATE, string RD_DATE, string PROD_PRICE, string PO_ID)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "UPDATE ZTBPOTable SET CTR_CD = '002', CUST_CD='" + CUST_CD + "', EMPL_NO='" + EMPL_NO + "', G_CODE = '" + G_CODE + "', PO_NO='" + PO_NO +"', PO_QTY='"+ PO_QTY + "', PO_DATE='" + PO_DATE + "', RD_DATE= '"+ RD_DATE + "', PROD_PRICE='" + PROD_PRICE + "'  WHERE PO_ID=" + PO_ID;
            result = config.GetData(strQuery);
            return result;

        }



        public DataTable InsertPO(string CTR_CD, string CUST_CD, string EMPL_NO, string G_CODE, string PO_NO, string PO_QTY, string PO_DATE, string RD_DATE, string PROD_PRICE)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "INSERT INTO ZTBPOTable (CTR_CD,CUST_CD,EMPL_NO,G_CODE,PO_NO,PO_QTY,PO_DATE,RD_DATE,PROD_PRICE) VALUES ('" + CTR_CD + "','" + CUST_CD + "','" + EMPL_NO + "','" + G_CODE + "','" + PO_NO + "','" + PO_QTY + "','" + PO_DATE + "','" + RD_DATE + "','" + PROD_PRICE + "')";
            result = config.GetData(strQuery);
            return result;

        }
        public DataTable InsertFCST(string CTR_CD, string EMPL_NO, string CUST_CD, string G_CODE, string PROD_PRICE, string YEAR, string WEEKNO, string W1, string W2, string W3, string W4, string W5, string W6, string W7, string W8, string W9, string W10, string W11, string W12, string W13, string W14, string W15, string W16, string W17, string W18, string W19, string W20, string W21, string W22)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "INSERT INTO ZTBFCSTTB (CTR_CD,EMPL_NO,CUST_CD,G_CODE,PROD_PRICE,FCSTYEAR,FCSTWEEKNO,W1,W2,W3,W4,W5,W6,W7,W8,W9,W10,W11,W12,W13,W14,W15,W16,W17,W18,W19,W20,W21,W22) VALUES ('" + CTR_CD + "','" + EMPL_NO + "','" + CUST_CD + "','" + G_CODE + "','" + PROD_PRICE + "','" + YEAR + "','" + WEEKNO + "','" + W1 + "','"  + W2 +  "','" + W3 + "','"  + W4 + "','" +  W5 + "','"  + W6 + "','"  + W7 + "','"  + W8 + "','" + W9 + "','" +  W10 + "','"  + W11 + "','"  + W12 + "','"  + W13 + "','"  + W14 + "','"  + W15 + "','"  + W16 + "','" + W17 + "','"  + W18 + "','"  + W19 + "','"  + W20 + "','" + W21 + "','" + W22 + "')";
           // MessageBox.Show(strQuery);
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable InsertPlan(string CTR_CD, string EMPL_NO, string CUST_CD, string G_CODE, string PLAN_DATE, string D1, string D2, string D3, string D4, string D5, string D6, string D7, string D8, string REMARK)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            //MessageBox.Show(D8);
            string strQuery = "INSERT INTO ZTBPLANTB (CTR_CD,EMPL_NO,CUST_CD,G_CODE,PLAN_DATE,D1,D2,D3,D4,D5,D6,D7,D8,REMARK) VALUES ('" + CTR_CD + "','" + EMPL_NO + "','" + CUST_CD + "','" + G_CODE + "','" + PLAN_DATE + "','" + D1 + "','" + D2 + "','" + D3 + "','" + D4 + "','" + D5 + "','" + D6 + "','" + D7 + "','" + D8 + "','" + REMARK +  "')";
            result = config.GetData(strQuery);
            return result;

        }


        public DataTable InsertYCSX(string CTR_CD, string PROD_REQUEST_DATE, string PROD_REQUEST_NO, string CODE_50, string CODE_03, string CODE_55, string G_CODE, string RIV_NO,string PROD_REQUEST_QTY, string CUST_CD, string EMPL_NO, string REMK, string INS_EMPL, string UPD_EMPL, string DELIVERY_DT, string PO, string TP, string BTP, string CK, string TOTAL_FCST, string W1, string W2, string W3, string W4, string W5, string W6, string W7, string W8, string PDUYET, string BLOCK_QTY)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "INSERT INTO P400 (CTR_CD,PROD_REQUEST_DATE,PROD_REQUEST_NO,CODE_50,CODE_03,CODE_55,G_CODE,RIV_NO,PROD_REQUEST_QTY,CUST_CD,EMPL_NO,REMK,INS_EMPL,UPD_EMPL, DELIVERY_DT, G_CODE2, PO_TDYCSX, TKHO_TDYCSX, BTP_TDYCSX, CK_TDYCSX, FCST_TDYCSX, W1, W2, W3, W4, W5, W6, W7, W8, PDUYET, BLOCK_TDYCSX) VALUES ('" + CTR_CD + "','"+ PROD_REQUEST_DATE+ "','" + PROD_REQUEST_NO + "','" + CODE_50 + "','" + CODE_03 + "','" + CODE_55 + "','" + G_CODE + "','" + RIV_NO + "','" + PROD_REQUEST_QTY + "','" + CUST_CD + "','" + EMPL_NO + "','" + REMK + "','" + INS_EMPL + "','" + UPD_EMPL + "','" + DELIVERY_DT + "','" + G_CODE+ "','" + PO + "','" + TP + "','" + BTP + "','" + CK + "','" + TOTAL_FCST + "','" + W1 + "','" + W2 + "','" + W3 + "','" + W4 + "','" + W5 + "','" + W6 + "','" + W7 + "','" + W8 + "' ,'" + PDUYET + "','" + BLOCK_QTY + "')";
            //MessageBox.Show(strQuery);
            result = config.GetData(strQuery);
            
            return result;

        }

       
        public DataTable insertM300(string CTR_CD, string OUT_DATE, string OUT_NO, string CODE_03, string CODE_50, string CODE_52, string DEPT_CD, string PROD_REQUEST_DATE, string PROD_REQUEST_NO, string INS_EMPL, string UPD_EMPL)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "INSERT INTO M300 (CTR_CD, OUT_DATE, OUT_NO, CODE_03, CODE_50, CODE_52, DEPT_CD, PROD_REQUEST_DATE, PROD_REQUEST_NO, INS_EMPL, UPD_EMPL) VALUES('"+CTR_CD+"','"+ OUT_DATE + "','" + OUT_NO + "','" + CODE_03 + "','" + CODE_50 + "','" + CODE_52 + "','" + DEPT_CD + "','" + PROD_REQUEST_DATE + "','" + PROD_REQUEST_NO + "','" + INS_EMPL + "','" + UPD_EMPL + "')";
            result = config.GetData(strQuery);
            return result;

        }
        public DataTable insertM301(string CTR_CD, string OUT_DATE, string OUT_NO, string OUT_SEQ, string CODE_03, string M_CODE, string OUT_PRE_QTY, string OUT_CFM_QTY ,string INS_EMPL, string UPD_EMPL)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "INSERT INTO M301 (CTR_CD, OUT_DATE, OUT_NO, OUT_SEQ, CODE_03, M_CODE, OUT_PRE_QTY, OUT_CFM_QTY, INS_EMPL, UPD_EMPL) VALUES ('" + CTR_CD + "','" + OUT_DATE + "','" + OUT_NO + "','" + OUT_SEQ + CODE_03 + "','" + M_CODE + "','" + OUT_PRE_QTY + "','" + OUT_CFM_QTY + "','" + INS_EMPL + "','" + UPD_EMPL +  "')";
            result = config.GetData(strQuery);
            return result;

        }



        public DataTable downLoadYCSX(string YCSXNO)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "Select CTR_CD,PROD_REQUEST_DATE,PROD_REQUEST_NO,CODE_50,CODE_03,CODE_55,G_CODE,RIV_NO,PROD_REQUEST_QTY,CUST_CD,EMPL_NO,REMK,INS_EMPL,UPD_EMPL,DELIVERY_DT FROM P400 WHERE PROD_REQUEST_NO='" + YCSXNO + "'";
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable getEmployeeName(string keyword)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "Select EMPL_NAME From M010  WHERE EMPL_NO='" + keyword + "'";
            result = config.GetData(strQuery);
            return result;

        }

        



        public DataTable login(string user, string pass)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "Select EMPL_NAME From M010  WHERE EMPL_NO='" + user + "' AND PASSWD='"+pass+"'";
            result = config.GetData(strQuery);
            return result;

        }


        public DataTable getYCSXInfo(string keyword)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "Select G_CODE,PROD_REQUEST_QTY, PROD_REQUEST_DATE,PROD_REQUEST_NO,REMK,DELIVERY_DT,CODE_50,CODE_55 From P400  WHERE PROD_REQUEST_NO='" + keyword + "'";
            result = config.GetData(strQuery);
            return result;

        }
        public DataTable getYCSXInfo2(string ycsxno)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = $"SELECT P400.CODE_50, P400.G_CODE, M100.G_NAME, P400.PROD_REQUEST_DATE, P400.PROD_REQUEST_NO, P400.PROD_REQUEST_QTY FROM P400 JOIN M100 ON (P400.G_CODE = M100.G_CODE) WHERE P400.PROD_REQUEST_NO='{ycsxno}'";
            //MessageBox.Show(strQuery);
            result = config.GetData(strQuery);
            return result;
        }
        public DataTable getcavity_print(string g_code)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = $"SELECT TOP 1 CAVITY_PRINT FROM BOM_AMAZONE LEFT JOIN DESIGN_AMAZONE ON (BOM_AMAZONE.G_CODE_MAU = DESIGN_AMAZONE.G_CODE_MAU AND  BOM_AMAZONE.DOITUONG_NO = DESIGN_AMAZONE.DOITUONG_NO) WHERE BOM_AMAZONE.G_CODE='{g_code}'";
            //MessageBox.Show(strQuery);
            result = config.GetData(strQuery);
            return result;
        }

        public DataTable getProductInfo(string keyword)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "Select G_NAME,G_CODE,G_WIDTH,G_LENGTH From M100";
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable testQuery(string keyword)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = keyword;
            result = config.GetData(strQuery);
            return result;

        }
        public DataTable checkOutMaterial(string keyword)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = "SELECT O302.OUT_DATE, O302.OUT_NO, M100.G_NAME,P400.G_CODE,O302.M_CODE, O302.M_LOT_NO,M090.M_NAME, M090.WIDTH_CD, O302.OUT_CFM_QTY, O300.PROD_REQUEST_DATE, O300.PROD_REQUEST_NO FROM O302 LEFT JOIN O300 ON (O300.OUT_DATE = O302.OUT_DATE AND O300.OUT_NO = O302.OUT_NO) LEFT JOIN M090 ON(O302.M_CODE = M090.M_CODE) LEFT JOIN P400 ON (P400.PROD_REQUEST_NO = O300.PROD_REQUEST_NO)  LEFT JOIN M100 ON (P400.G_CODE = M100.G_CODE)  WHERE O302.M_LOT_NO IN" + keyword;
            //string strQuery = keyword;
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable checkLastYCSX(string keyword)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "Select M100.G_NAME, P400.CTR_CD, MAX(P400.PROD_REQUEST_DATE) AS PROD_REQUEST_DATE,P400.PROD_REQUEST_NO,P400.CODE_50,P400.CODE_03,P400.CODE_55,P400.G_CODE,P400.RIV_NO,P400.PROD_REQUEST_QTY,P400.CUST_CD,P400.EMPL_NO,P400.REMK,P400.INS_EMPL,P400.UPD_EMPL,P400.DELIVERY_DT,CASE WHEN CODE_50 = '01' THEN 'GC' WHEN CODE_50 = '02' THEN 'SK' WHEN CODE_50 = '03' THEN 'KD' WHEN CODE_50 = '04' THEN 'VN' WHEN CODE_50 = '05' THEN 'Sample' WHEN CODE_50 = '06' THEN 'Vai bac 4' WHEN CODE_50 = '07' THEN 'ETC' END AS PHAN_LOAI_NHAP_KHAU, CASE WHEN CODE_55 = '01' THEN 'thong thuong' WHEN CODE_55 = '02' THEN 'sdi' WHEN CODE_55 = '03' THEN 'etc' WHEN CODE_55 = '04' THEN 'Sample' END AS PHAN_LOAI_SAN_XUAT  FROM P400 LEFT JOIN M100 ON P400.G_CODE = M100.G_CODE WHERE M100.G_NAME LIKE '%GH63-17671A%'";            
            string strQuery = "SELECT P400.EMPL_NO, P400.PROD_REQUEST_NO, P400.PROD_REQUEST_DATE,P400.PROD_REQUEST_QTY,TB2.G_CODE, TB2.G_NAME FROM P400 INNER JOIN (SELECT M100.G_NAME, NEWTB.G_CODE, NEWTB.MAX_PROD_REQUEST_DATE FROM M100 INNER JOIN (SELECT MAX(P400.PROD_REQUEST_DATE) AS MAX_PROD_REQUEST_DATE, P400.G_CODE FROM P400  WHERE G_CODE IN (SELECT M100.G_CODE FROM M100 WHERE  " + keyword + ") GROUP BY P400.G_CODE) AS NEWTB ON NEWTB.G_CODE = M100.G_CODE) AS TB2 ON (TB2.G_CODE = P400.G_CODE AND TB2.MAX_PROD_REQUEST_DATE = P400.PROD_REQUEST_DATE) ";
            result = config.GetData(strQuery);
            return result;
        }

        public DataTable getnameandcode(string keyword)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "Select G_NAME,G_CODE From M100 WHERE G_NAME LIKE'%" + keyword + "%'";
            result = config.GetData(strQuery);
            return result;
        }

        public DataTable getamazonedesign()
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = $"SELECT DISTINCT M100.G_NAME, M100.G_CODE FROM DESIGN_AMAZONE LEFT JOIN M100 ON (M100.G_CODE = DESIGN_AMAZONE.G_CODE_MAU)";
            result = config.GetData(strQuery);
            return result;
        }
        public DataTable getamazonedesign2(string G_CODE)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = $"SELECT DESIGN_AMAZONE.G_CODE_MAU, M100.G_NAME AS TEN_MAU, DOITUONG_NO, DOITUONG_NAME FROM DESIGN_AMAZONE JOIN M100 ON (M100.G_CODE = DESIGN_AMAZONE.G_CODE_MAU) WHERE G_CODE_MAU ='{G_CODE}'";
            result = config.GetData(strQuery);
            return result;
        }



        public DataTable getcustomerinfo(string keyword)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "Select CUST_CD,CUST_NAME From M110 WHERE CUST_NAME_KD LIKE'%" + keyword + "%'";
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable getProductBOM(string keyword)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "Select M140.M_CODE,M140.M_QTY,M140.RIV_NO,M090.M_NAME,M090.WIDTH_CD From M140  LEFT JOIN M090 ON M140.M_CODE = M090.M_CODE WHERE G_CODE='" + keyword + "'";
            result = config.GetData(strQuery);
            return result;

        }
        public DataTable getFullInfo(string keyword)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = "SELECT P400.REMK,P400.PROD_REQUEST_QTY,P400.PROD_REQUEST_NO,P400.PROD_REQUEST_DATE,P400.G_CODE,P400.DELIVERY_DT,P400.CODE_55,P400.CODE_50,M140.RIV_NO,M140.M_QTY,M140.M_CODE,M110.CUST_NAME,M100.ROLE_EA_QTY,M100.PACK_DRT,M100.G_WIDTH,M100.G_SG_R,M100.G_SG_L,M100.G_R,M100.G_NAME,M100.G_LG,M100.G_LENGTH,M100.G_CODE_C,M100.G_CG,M100.G_C,M100.CODE_33,M090.M_NAME,M090.WIDTH_CD,M010.EMPL_NO,M010.EMPL_NAME, P400.CODE_03,M140.REMK AS REMARK  FROM P400 LEFT JOIN M100  ON P400.G_CODE = M100.G_CODE  LEFT JOIN M010 ON P400.EMPL_NO = M010.EMPL_NO  RIGHT JOIN M140 ON P400.G_CODE = M140.G_CODE LEFT JOIN M090 ON M090.M_CODE = M140.M_CODE LEFT JOIN M110 ON M110.CUST_CD = P400.CUST_CD    WHERE P400.PROD_REQUEST_NO = '" + keyword + "'";
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable getProductionInfo()
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = "SELECT AA.EQUIPMENT_CD, AA.LAST_ACTIVE_TIME, M100.G_CODE, M100.G_NAME, M100.G_NAME_KD, M090.M_NAME, M110.CUST_NAME_KD, P500.PROD_REQUEST_NO, M010.EMPL_NAME FROM (SELECT DISTINCT EQUIPMENT_CD,  MAX(INS_DATE) OVER (PARTITION BY EQUIPMENT_CD) AS LAST_ACTIVE_TIME FROM P500) AS AA LEFT JOIN P500 ON (AA.LAST_ACTIVE_TIME = P500.INS_DATE) LEFT JOIN M100 ON (P500.G_CODE = M100.G_CODE) LEFT JOIN M090 ON (M090.M_CODE = P500.M_CODE) LEFT JOIN P400 ON (P400.PROD_REQUEST_NO = P500.PROD_REQUEST_NO AND P400.PROD_REQUEST_DATE = P500.PROD_REQUEST_DATE) LEFT JOIN M010 ON (P400.EMPL_NO = M010.EMPL_NO) LEFT JOIN M110 ON (M110.CUST_CD = P400.CUST_CD) ORDER BY  AA.LAST_ACTIVE_TIME DESC";
            result = config.GetData(strQuery);
            return result;

        }
        public DataTable getConfig(string prod_type, string size, string eq, int step)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = $"SELECT * FROM  ZTB_TBG_CONFIG WHERE (PROD_TYPE='{prod_type}'AND EQ='{eq}' AND SIZE='{size}' AND STEP={step})";
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable getBEPConfig(string eq)
        { 
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = $"SELECT * FROM  ZTB_BEP_CONFIG WHERE EQ='{eq}'";
            result = config.GetData(strQuery);
            return result;

        }



        public DataTable getFullInfoBOM2(string keyword)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = $"SELECT ZTB_BOM2.BOM_ID AS ZTB_BOM2_BOM_ID,ZTB_BOM2.G_CODE AS ZTB_BOM2_G_CODE,ZTB_BOM2.RIV_NO AS ZTB_BOM2_RIV_NO,ZTB_BOM2.G_SEQ AS ZTB_BOM2_G_SEQ,ZTB_BOM2.CATEGORY AS ZTB_BOM2_CATEGORY,ZTB_BOM2.M_CODE AS ZTB_BOM2_M_CODE,ZTB_BOM2.M_NAME AS ZTB_BOM2_M_NAME,ZTB_BOM2.CUST_CD AS ZTB_BOM2_CUST_CD,ZTB_BOM2.IMPORT_CAT AS ZTB_BOM2_IMPORT_CAT,ZTB_BOM2.M_CMS_PRICE AS ZTB_BOM2_M_CMS_PRICE,ZTB_BOM2.M_SS_PRICE AS ZTB_BOM2_M_SS_PRICE,ZTB_BOM2.M_SLITTING_PRICE AS ZTB_BOM2_M_SLITTING_PRICE,ZTB_BOM2.USAGE AS ZTB_BOM2_USAGE,ZTB_BOM2.MAT_MASTER_WIDTH AS ZTB_BOM2_MAT_MASTER_WIDTH,ZTB_BOM2.MAT_CUTWIDTH AS ZTB_BOM2_MAT_CUTWIDTH,ZTB_BOM2.MAT_ROLL_LENGTH AS ZTB_BOM2_MAT_ROLL_LENGTH,ZTB_BOM2.MAT_THICKNESS AS ZTB_BOM2_MAT_THICKNESS,ZTB_BOM2.M_QTY AS ZTB_BOM2_M_QTY,ZTB_BOM2.REMARK AS ZTB_BOM2_REMARK,ZTB_BOM2.PROCESS_ORDER AS ZTB_BOM2_PROCESS_ORDER,ZTB_BOM2.INS_EMPL AS ZTB_BOM2_INS_EMPL,ZTB_BOM2.UPD_EMPL AS ZTB_BOM2_UPD_EMPL,ZTB_BOM2.INS_DATE AS ZTB_BOM2_INS_DATE,ZTB_BOM2.UPD_DATE AS ZTB_BOM2_UPD_DATE, M100.G_CODE AS M100_G_CODE,M100.G_NAME AS M100_G_NAME,M100.CODE_12 AS M100_CODE_12,M100.SEQ_NO AS M100_SEQ_NO,M100.REV_NO AS M100_REV_NO,M100.CODE_33 AS M100_CODE_33,M100.CUST_CD AS M100_CUST_CD,M100.G_CODE_C AS M100_G_CODE_C,M100.G_CODE_V AS M100_G_CODE_V,M100.G_CODE_K AS M100_G_CODE_K,M100.USD_AMT AS M100_USD_AMT,M100.VND_AMT AS M100_VND_AMT,M100.KRW_AMT AS M100_KRW_AMT,M100.CODE_27 AS M100_CODE_27,M100.CODE_28 AS M100_CODE_28,M100.PRT_DRT AS M100_PRT_DRT,M100.PRT_YN AS M100_PRT_YN,M100.CODE_32 AS M100_CODE_32,M100.PACK_DRT AS M100_PACK_DRT,M100.ROLE_EA_QTY AS M100_ROLE_EA_QTY,M100.G_WIDTH AS M100_G_WIDTH,M100.G_LENGTH AS M100_G_LENGTH,M100.G_R AS M100_G_R,M100.G_C AS M100_G_C,M100.G_LG AS M100_G_LG,M100.G_SG_L AS M100_G_SG_L,M100.G_SG_R AS M100_G_SG_R,M100.G_CG AS M100_G_CG,M100.META_PAT_CD AS M100_META_PAT_CD,M100.CODE_34 AS M100_CODE_34,M100.CODE_35 AS M100_CODE_35,M100.RIBON_SPEC AS M100_RIBON_SPEC,M100.CODE_36 AS M100_CODE_36,M100.REMK AS M100_REMK,M100.USE_YN AS M100_USE_YN,M100.INS_DATE AS M100_INS_DATE,M100.INS_EMPL AS M100_INS_EMPL,M100.UPD_DATE AS M100_UPD_DATE,M100.UPD_EMPL AS M100_UPD_EMPL,M100.PROD_TYPE AS M100_PROD_TYPE,M100.PROD_MODEL AS M100_PROD_MODEL,M100.PROD_DIECUT_STEP AS M100_PROD_DIECUT_STEP,M100.PROD_PRINT_TIMES AS M100_PROD_PRINT_TIMES,M100.PROD_MAIN_MATERIAL AS M100_PROD_MAIN_MATERIAL,M100.PROD_PROJECT AS M100_PROD_PROJECT,M100.G_NAME_KD AS M100_G_NAME_KD,M100.PROD_LAST_PRICE AS M100_PROD_LAST_PRICE,M100.PROD_LAST_PRICE_UPDATED_DATE AS M100_PROD_LAST_PRICE_UPDATED_DATE,M100.FACTORY AS M100_FACTORY,M100.EQ1 AS M100_EQ1,M100.EQ2 AS M100_EQ2,M100.Setting1 AS M100_Setting1,M100.UPH1 AS M100_UPH1,M100.Step1 AS M100_Step1,M100.Setting2 AS M100_Setting2,M100.UPH2 AS M100_UPH2,M100.Step2 AS M100_Step2,M100.NOTE AS M100_NOTE,M100.INSPECT_SPEED AS M100_INSPECT_SPEED,M100.DESCR AS M100_DESCR,M100.DRAW_LINK AS M100_DRAW_LINK,M100.PD AS M100_PD,M100.G_C_R AS M100_G_C_R,M100.KNIFE_TYPE AS M100_KNIFE_TYPE,M100.KNIFE_LIFECYCLE AS M100_KNIFE_LIFECYCLE,M100.RPM AS M100_RPM,M100.PIN_DISTANCE AS M100_PIN_DISTANCE,M100.PROCESS_TYPE AS M100_PROCESS_TYPE,M100.PACKING_FEE AS M100_PACKING_FEE,M100.PROFIT_RATE AS M100_PROFIT_RATE,M100.PROD_MANPOWER AS M100_PROD_MANPOWER,M100.INSPECT_MANPOWER AS M100_INSPECT_MANPOWER,M100.KNIFE_PRICE AS M100_KNIFE_PRICE,M100.PRODUCT_CMSPRICE AS M100_PRODUCT_CMSPRICE,M100.PRODUCT_SSPRICE AS M100_PRODUCT_SSPRICE,M100.PRODUCT_FINAL_PRICE AS M100_PRODUCT_FINAL_PRICE, M110.CUST_NAME_KD AS M110_CUST_NAME_KD  FROM ZTB_BOM2 JOIN M100 ON (ZTB_BOM2.G_CODE = M100.G_CODE) JOIN M110 ON (M110.CUST_CD = M100.CUST_CD) WHERE M100.G_CODE='{keyword}'";
            result = config.GetData(strQuery);
           // MessageBox.Show(strQuery);
            return result;

        }
        public DataTable getFullBOM(string G_CODE, string RIV_NO )
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = $"SELECT M140.G_CODE, M100.G_NAME, M100.G_NAME_KD, M140.RIV_NO, M140.M_CODE, M090.M_NAME, M090.WIDTH_CD, M140.M_QTY, M140.INS_EMPL, M140.INS_DATE, M140.UPD_EMPL,M140.UPD_DATE FROM M140 JOIN M100 ON (M140.G_CODE = M100.G_CODE) JOIN M090 ON (M090.M_CODE = M140.M_CODE) WHERE M140.G_CODE='{G_CODE}' AND M140.RIV_NO='{RIV_NO}'";
            //MessageBox.Show(strQuery);
            result = config.GetData(strQuery);
            return result;
        }
        public DataTable getFullBOMXuatLieu(string G_CODE, string RIV_NO)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = $"SELECT  M140.RIV_NO, M140.M_CODE, M090.M_NAME, M090.WIDTH_CD FROM M140 JOIN M100 ON (M140.G_CODE = M100.G_CODE) JOIN M090 ON (M090.M_CODE = M140.M_CODE) WHERE M140.G_CODE='{G_CODE}' AND M140.RIV_NO='{RIV_NO}'";
            //MessageBox.Show(strQuery);
            result = config.GetData(strQuery);
            return result;
        }
        public DataTable getFullBOM2(string G_CODE, string RIV_NO)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = $"SELECT * FROM ZTB_BOM2 WHERE G_CODE='{G_CODE}' AND RIV_NO='{RIV_NO}'";
            //MessageBox.Show(strQuery);
            result = config.GetData(strQuery);
            return result;
        }

        public DataTable getbeplist(string G_CODE, string chuabep)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = $"SELECT AA.G_CODE, M100.PROD_TYPE, G_NAME, G_NAME_KD, EQ1, EQ2, PD, (G_C * G_C_R) AS CAVITY, PROD_MANPOWER, INSPECT_MANPOWER, M100.BEP_1HOUR_PROD_QTY, M100.BEP_PROD_NG_RATE,  M100.BEP_INSP_NG_RATE FROM (SELECT DISTINCT G_CODE FROM ZTB_BOM2) AS AA JOIN M100 ON (AA.G_CODE = M100.G_CODE) JOIN M110 ON (M110.CUST_CD = M100.CUST_CD) WHERE (AA.G_CODE ='{G_CODE}' OR G_NAME LIKE '%{G_CODE}%')";
            if (chuabep == "chuabep") strQuery += " AND BEP_INSP_NG_RATE is null";
            result = config.GetData(strQuery);
            return result;
        }



        public DataTable getMaterialInfo(string m_name, string thieuinfo)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = $"SELECT M_ID,M_NAME,CUST_CD,SSPRICE,CMSPRICE,SLITTING_PRICE,MASTER_WIDTH,ROLL_LENGTH FROM ZTB_MATERIAL_TB WHERE M_NAME LIKE '%{m_name}%'";
            if(thieuinfo == "thieuinfo")
            {
                strQuery += " AND (CUST_CD is null OR SSPRICE is null OR CMSPRICE is null OR SLITTING_PRICE is null OR MASTER_WIDTH is null OR ROLL_LENGTH is null)";
            }
            strQuery += "AND 1=1 ORDER BY M_ID DESC";
            result = config.GetData(strQuery);
            return result;
        }
        public DataTable getBOM2Info(string g_name, string onlyLieu)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery =$"SELECT  M100.G_NAME, ZTB_BOM2.BOM_ID,ZTB_BOM2.G_CODE,ZTB_BOM2.CATEGORY,ZTB_BOM2.M_CODE,ZTB_BOM2.M_NAME,ZTB_BOM2.CUST_CD,ZTB_BOM2.IMPORT_CAT,ZTB_BOM2.M_CMS_PRICE,ZTB_BOM2.M_SS_PRICE,ZTB_BOM2.M_SLITTING_PRICE,ZTB_BOM2.USAGE,ZTB_BOM2.MAT_MASTER_WIDTH,ZTB_BOM2.MAT_CUTWIDTH,ZTB_BOM2.MAT_ROLL_LENGTH,ZTB_BOM2.MAT_THICKNESS,ZTB_BOM2.M_QTY,ZTB_BOM2.REMARK,ZTB_BOM2.PROCESS_ORDER FROM ZTB_BOM2 JOIN M100 ON (ZTB_BOM2.G_CODE = M100.G_CODE) WHERE (M100.G_NAME LIKE '%{g_name}%' OR M100.G_CODE='{g_name}')  ";

            if(onlyLieu == "lieu") {
                strQuery += "AND ZTB_BOM2.CATEGORY=1";
            }
            result = config.GetData(strQuery);
            return result;
        }

        public DataTable getgiafullBOM2Info(string g_name, string onlyLieu)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = $"SELECT  M100.G_NAME, ZTB_BOM2.BOM_ID,ZTB_BOM2.G_CODE,ZTB_BOM2.CATEGORY,ZTB_BOM2.M_CODE,ZTB_BOM2.M_NAME,ZTB_BOM2.CUST_CD,ZTB_BOM2.IMPORT_CAT,ZTB_BOM2.M_CMS_PRICE,ZTB_BOM2.M_SS_PRICE,ZTB_BOM2.M_SLITTING_PRICE,ZTB_BOM2.USAGE,ZTB_BOM2.MAT_MASTER_WIDTH,ZTB_BOM2.MAT_CUTWIDTH,ZTB_BOM2.MAT_ROLL_LENGTH,ZTB_BOM2.MAT_THICKNESS,ZTB_BOM2.M_QTY,ZTB_BOM2.REMARK,ZTB_BOM2.PROCESS_ORDER, M110.CUST_NAME_KD,M100.G_NAME_KD,M100.G_WIDTH,M100.G_LENGTH,M100.G_C,M100.G_C_R,(M100.G_C*M100.G_C_R) AS CAVITY,M100.PD,M100.PROD_TYPE,M100.PROD_MODEL,M100.PROD_DIECUT_STEP,M100.PROD_PRINT_TIMES,M100.PROD_MAIN_MATERIAL,M100.PROD_PROJECT,M100.EQ1,M100.EQ2,M100.DESCR,M100.DRAW_LINK,M100.KNIFE_TYPE,M100.KNIFE_LIFECYCLE,M100.RPM,M100.PIN_DISTANCE,M100.PROCESS_TYPE,M100.KNIFE_PRICE, M100.REMK,M100.USE_YN, M100.PROD_MANPOWER, M100.INSPECT_MANPOWER, M100.BEP_1HOUR_PROD_QTY, M100.BEP_PROD_NG_RATE,  M100.BEP_INSP_NG_RATE,M100.MATERIAL_COST_CMS,M100.PROCESS_COST_CMS,M100.OTHER_COST_CMS,M100.PROFIT_VALUE_CMS,M100.MATERIAL_COST_SS,M100.PROCESS_COST_SS,M100.OTHER_COST_SS,M100.PROFIT_VALUE_SS,M100.PRODUCT_CMSPRICE,M100.MCR_CMS,M100.PRODUCT_SSPRICE,M100.MCR_SS, M100.BEP_MAT_COST, M100.BEP_PROC_COST, M100.BEP_TOTAL_LOSS, M100.BEP_PROFIT_VALUE, M100.BEP_PRICE, M100.BEP_TARGET_PRICE, M100.PRODUCT_FINAL_PRICE FROM ZTB_BOM2 JOIN M100 ON (ZTB_BOM2.G_CODE = M100.G_CODE) JOIN M110 ON (M110.CUST_CD = M100.CUST_CD) WHERE (M100.G_NAME LIKE '%{g_name}%' OR M100.G_CODE='{g_name}')  ";

            if (onlyLieu == "lieu")
            {
                strQuery += "AND ZTB_BOM2.CATEGORY=1";
            }
            result = config.GetData(strQuery);
            return result;
        }


        public DataTable updateBEPInfo(string updatevalue)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = $"UPDATE M100 " + updatevalue;
            result = config.GetData(strQuery);
            return result;
        }

        public DataTable updateMaterial(string updatevalue)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = $"UPDATE ZTB_MATERIAL_TB " + updatevalue;
            result = config.GetData(strQuery);
            return result;
        }
        public DataTable updateBOM2(string updatevalue)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = $"UPDATE ZTB_BOM2 " + updatevalue;
            result = config.GetData(strQuery);
            return result;
        }

        public DataTable updatebaogiaM100(string updatevalue)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = $"UPDATE M100 " + updatevalue;
            result = config.GetData(strQuery);
            return result;
        }

        public DataTable updateConfig(string updatevalue)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = $"UPDATE ZTB_TBG_CONFIG " + updatevalue;
            result = config.GetData(strQuery);
            return result;
        }


        public DataTable insertMaterialfromBOMtoMTable()
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = $"INSERT INTO ZTB_MATERIAL_TB  (CTR_CD,M_NAME) SELECT DISTINCT '002', ZTB_BOM2.M_NAME FROM ZTB_BOM2 WHERE (ZTB_BOM2.M_NAME not in (SELECT M_NAME FROM  ZTB_MATERIAL_TB) AND ZTB_BOM2.CATEGORY = 1)";
            result = config.GetData(strQuery);
            return result;
        }

        public DataTable checkDataAmazone(string DATA)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            int kq = 0;
            string strQuery = $"SELECT * FROM AMAZONE_DATA WHERE (DATA_1='{DATA}' OR DATA_2='{DATA}' OR DATA_3='{DATA}' OR DATA_4='{DATA}')";
            //MessageBox.Show(strQuery);
            result = config.GetData(strQuery);
            return result;
        }
        public DataTable checkAMZOriginalCount()
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            int kq = 0;
            string strQuery = $"SELECT count(AA.code) as originalcodecount FROM (SELECT code FROM ( SELECT DATA_1 FROM AMAZONE_DATA UNION ALL SELECT DATA_2 FROM AMAZONE_DATA ) AS UNIQUEDATA(code) ) as AA";
            //MessageBox.Show(strQuery);
            result = config.GetData(strQuery);
            return result;
        }
        public DataTable checkAMZUniqueCount()
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            int kq = 0;
            string strQuery = $"SELECT count(distinct AA.code) as uniquecodecount FROM (SELECT code FROM ( SELECT DATA_1 FROM AMAZONE_DATA UNION ALL SELECT DATA_2 FROM AMAZONE_DATA ) AS UNIQUEDATA(code) ) as AA";
            //MessageBox.Show(strQuery);
            result = config.GetData(strQuery);
            return result;
        }
        public DataTable checkAMZDuplicateCount()
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            int kq = 0;
            string strQuery = $"SELECT * FROM AMAZONE_DATA WHERE DATA_1 IN( SELECT AAA.originalcodecount FROM ( SELECT AA.code as originalcodecount, COUNT(AA.code) AS COUNT_ FROM (SELECT code FROM ( SELECT DATA_1 FROM AMAZONE_DATA UNION ALL SELECT DATA_2 FROM AMAZONE_DATA ) AS UNIQUEDATA(code) ) as AA GROUP BY AA.code HAVING COUNT(AA.code) >1 ) AS AAA ) UNION SELECT * FROM AMAZONE_DATA WHERE DATA_2 IN( SELECT AAA.originalcodecount FROM ( SELECT AA.code as originalcodecount, COUNT(AA.code) AS COUNT_ FROM (SELECT code FROM ( SELECT DATA_1 FROM AMAZONE_DATA UNION ALL SELECT DATA_2 FROM AMAZONE_DATA ) AS UNIQUEDATA(code) ) as AA GROUP BY AA.code HAVING COUNT(AA.code) >1 ) AS AAA )";
            //MessageBox.Show(strQuery);
            result = config.GetData(strQuery);
            return result;
        }
        public DataTable checkModelNameAmazone(string G_CODE)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            int kq = 0;
            string strQuery = $"SELECT PROD_MODEL FROM M100 WHERE G_CODE='{G_CODE}'";
            //MessageBox.Show(strQuery);
            result = config.GetData(strQuery);
            return result;
        }


        public DataTable checkNO_IN_Amazone(string NO_IN)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            int kq = 0;
            string strQuery = $"SELECT * FROM AMAZONE_DATA WHERE NO_IN='{NO_IN}'";
            result = config.GetData(strQuery);
            return result;
        }


        public DataTable checkDESIGNAmazone(string G_CODE_MAU)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            int kq = 0;
            string strQuery = $"SELECT G_CODE_MAU,DOITUONG_NO,DOITUONG_NAME,PHANLOAI_DT,DOITUONG_STT,CAVITY_PRINT, GIATRI, FONT_NAME, FONT_SIZE,FONT_STYLE,POS_X,POS_Y,SIZE_W,SIZE_H,ROTATE,REMARK FROM  DESIGN_AMAZONE WHERE G_CODE_MAU='{G_CODE_MAU}'";
            result = config.GetData(strQuery);
            return result;
        }

        public DataTable checkBOMAmazone(string G_CODE)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            int kq = 0;
            string strQuery = $"SELECT BOM_AMAZONE.G_CODE, M100.G_NAME, DESIGN_AMAZONE.G_CODE_MAU,  M100_B.G_NAME AS TEN_MAU,BOM_AMAZONE.DOITUONG_NO, DESIGN_AMAZONE.DOITUONG_NAME, BOM_AMAZONE.GIATRI, BOM_AMAZONE.REMARK FROM BOM_AMAZONE LEFT JOIN DESIGN_AMAZONE ON (BOM_AMAZONE.G_CODE_MAU= DESIGN_AMAZONE.G_CODE_MAU AND BOM_AMAZONE.DOITUONG_NO= DESIGN_AMAZONE.DOITUONG_NO) LEFT JOIN M100 ON (M100.G_CODE = BOM_AMAZONE.G_CODE)  LEFT JOIN  (SELECT * FROM M100) AS M100_B ON (M100_B.G_CODE = DESIGN_AMAZONE.G_CODE_MAU) WHERE BOM_AMAZONE.G_CODE='{G_CODE}'";
            result = config.GetData(strQuery);           
            return result;
        }

        public string condition_amazone(string G_NAME, string NO_IN, string YCSX_NO)
        {
            string condition = " WHERE 1=1";
            if (G_NAME != "") condition += $" AND M100.G_NAME LIKE '%{G_NAME}%' ";
            if (NO_IN != "") condition += $" AND AMAZONE_DATA.NO_IN ='{NO_IN}' ";
            if (YCSX_NO != "") condition += $" AND AMAZONE_DATA.PROD_REQUEST_NO = '{YCSX_NO}' ";
            return condition;
        }
        public DataTable checkDATAAmazone(string G_NAME, string NO_IN, string YCSX_NO)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            int kq = 0;
            string strQuery = $"SELECT M100.G_NAME, AMAZONE_DATA.G_CODE, AMAZONE_DATA.PROD_REQUEST_NO, AMAZONE_DATA.NO_IN, AMAZONE_DATA.ROW_NO, AMAZONE_DATA.DATA_1, AMAZONE_DATA.DATA_2,AMAZONE_DATA.DATA_3,AMAZONE_DATA.DATA_4, AMAZONE_DATA.PRINT_STATUS, AMAZONE_DATA.INLAI_COUNT, AMAZONE_DATA.REMARK, AMAZONE_DATA.INS_DATE, AMAZONE_DATA.INS_EMPL FROM AMAZONE_DATA LEFT JOIN M100 ON (M100.G_CODE = AMAZONE_DATA.G_CODE)" + condition_amazone(G_NAME, NO_IN, YCSX_NO);
            result = config.GetData(strQuery);
            //MessageBox.Show(strQuery);
            return result;
        }


        public int checkBOM2(string G_CODE, string RIV_NO)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            int kq = 0;
            string strQuery = $"SELECT * FROM ZTB_BOM2 WHERE G_CODE='{G_CODE}' AND RIV_NO='{RIV_NO}'";
            result = config.GetData(strQuery);
            if(result.Rows.Count >0)
            {
                kq = 1;
            }
            return kq;
        }

        public int checkBOMSXExist(string G_CODE, string RIV_NO)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            int kq = 0;
            string strQuery = $"SELECT * FROM M140 WHERE G_CODE='{G_CODE}' AND RIV_NO='{RIV_NO}'";
            result = config.GetData(strQuery);
            if (result.Rows.Count > 0)
            {
                kq = 1;
            }
            return kq;
        }


        public DataTable insertAMAZONEDATA(string insertValue)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = $"INSERT INTO AMAZONE_DATA (CTR_CD,G_CODE,PROD_REQUEST_NO,NO_IN,ROW_NO,DATA_1,DATA_2,DATA_3,DATA_4,PRINT_STATUS,INLAI_COUNT,REMARK,INS_DATE,INS_EMPL,UPD_DATE,UPD_EMPL) VALUES " + insertValue;
            //MessageBox.Show(strQuery);
            result = config.GetData(strQuery);
            return result;

        }


        public DataTable insertOldBOM(string value)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = $"INSERT INTO M140 (CTR_CD,G_CODE,RIV_NO,G_SEQ,M_CODE,M_QTY,META_PAT_CD,REMK,USE_YN,INS_DATE,INS_EMPL,UPD_DATE,UPD_EMPL) VALUES " + value;           
            //MessageBox.Show(strQuery);
            result = config.GetData(strQuery);
            return result;

        }
        public DataTable checkProcessInNoP500(string in_date, string machine)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = $"SELECT TOP 1 PROCESS_IN_DATE, PROCESS_IN_NO, EQUIPMENT_CD FROM P500 WHERE PROCESS_IN_DATE='{in_date}'  ORDER BY INS_DATE DESC";
            //MessageBox.Show(strQuery);
            result = config.GetData(strQuery);
            return result;
        }

        public DataTable insertP500(string value)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = $"INSERT INTO P500 (CTR_CD, PROCESS_IN_DATE, PROCESS_IN_NO, PROCESS_IN_SEQ, M_LOT_IN_SEQ, PROD_REQUEST_DATE, PROD_REQUEST_NO, G_CODE, M_CODE, M_LOT_NO, EMPL_NO, EQUIPMENT_CD, SCAN_RESULT, INS_DATE, INS_EMPL, UPD_DATE, UPD_EMPL, FACTORY) VALUES " + value;
            //MessageBox.Show(strQuery);
            result = config.GetData(strQuery);
            return result;
        }
        public DataTable insertP501(string value)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = $"INSERT INTO P501 (CTR_CD,PROCESS_IN_DATE,PROCESS_IN_NO,PROCESS_IN_SEQ,M_LOT_IN_SEQ,PROCESS_PRT_SEQ,M_LOT_NO,PROCESS_LOT_NO,INS_DATE,INS_EMPL,UPD_DATE,UPD_EMPL) VALUES " + value;
            //MessageBox.Show(strQuery);
            result = config.GetData(strQuery);
            return result;
        }

        public DataTable insertNewBOM(string value)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = $"INSERT INTO ZTB_BOM2 (CTR_CD,G_CODE,RIV_NO,G_SEQ,CATEGORY,M_CODE,M_NAME,CUST_CD,IMPORT_CAT,M_CMS_PRICE,M_SS_PRICE,M_SLITTING_PRICE,USAGE,MAT_MASTER_WIDTH,MAT_CUTWIDTH,MAT_ROLL_LENGTH,MAT_THICKNESS,M_QTY,REMARK,PROCESS_ORDER,INS_DATE,INS_EMPL,UPD_DATE,UPD_EMPL) VALUES " + value;
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable deleteBOM(string g_code)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = $"DELETE FROM ZTB_BOM2 WHERE G_CODE = '{g_code}'";
            result = config.GetData(strQuery);
            return result;

        }
        public DataTable deleteBOMSX(string g_code)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = $"DELETE FROM M140 WHERE G_CODE = '{g_code}'";
            result = config.GetData(strQuery);
            return result;

        }


        public DataTable getMATInfo(string M_NAME)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = $"SELECT * FROM ZTB_MATERIAL_TB WHERE M_NAME ='{M_NAME}'";
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable getAmazoneBOM(string G_NAME)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = $"SELECT DISTINCT(M100.G_NAME), M100.G_CODE FROM BOM_AMAZONE JOIN M100 ON (M100.G_CODE = BOM_AMAZONE.G_CODE) WHERE M100.G_NAME LIKE '%{G_NAME}%'";
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable getAmazoneBOM_GCODE(string G_CODE)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = $"SELECT DISTINCT(M100.G_NAME), M100.G_CODE FROM BOM_AMAZONE JOIN M100 ON (M100.G_CODE = BOM_AMAZONE.G_CODE) WHERE M100.G_CODE = '{G_CODE}'";
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable getMassAmazone_GCODE(string G_CODE)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = $"SELECT * FROM P400 WHERE G_CODE= '{G_CODE}' AND CODE_55 <> '04'";
            result = config.GetData(strQuery);
            return result;

        }


        public DataTable getMcodeInfo(string keyword)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = $"SELECT * FROM M100 WHERE G_CODE LIKE '%{keyword}%' OR G_NAME LIKE '%{keyword}%'";
            result = config.GetData(strQuery);
            return result;
                  
        }

        public DataTable getcodephoiamazone(string keyword)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = $"SELECT M100.G_CODE, M100.G_NAME FROM M100 WHERE G_CODE LIKE '%{keyword}%' OR G_NAME LIKE '%{keyword}%'";
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable getBaoGiaConfig()
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = $"SELECT * FROM ZTB_TBG_CONFIG";
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable getcodebom2Info(string keyword, string tinhgiachua)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = $"SELECT M110.CUST_NAME_KD,M100.G_NAME_KD,M100.G_CODE,M100.G_NAME,M100.G_WIDTH,M100.G_LENGTH,M100.G_C,M100.G_C_R,(M100.G_C*M100.G_C_R) AS CAVITY,M100.PD,M100.PROD_TYPE,M100.PROD_MODEL,M100.PROD_DIECUT_STEP,M100.PROD_PRINT_TIMES,M100.PROD_MAIN_MATERIAL,M100.PROD_PROJECT,M100.EQ1,M100.EQ2,M100.DESCR,M100.DRAW_LINK,M100.KNIFE_TYPE,M100.KNIFE_LIFECYCLE,M100.RPM,M100.PIN_DISTANCE,M100.PROCESS_TYPE,M100.KNIFE_PRICE, M100.REMK,M100.USE_YN, M100.PROD_MANPOWER, M100.INSPECT_MANPOWER, M100.BEP_1HOUR_PROD_QTY, M100.BEP_PROD_NG_RATE,  M100.BEP_INSP_NG_RATE,M100.MATERIAL_COST_CMS,M100.PROCESS_COST_CMS,M100.OTHER_COST_CMS,M100.PROFIT_VALUE_CMS,M100.MATERIAL_COST_SS,M100.PROCESS_COST_SS,M100.OTHER_COST_SS,M100.PROFIT_VALUE_SS,M100.PRODUCT_CMSPRICE,M100.MCR_CMS,M100.PRODUCT_SSPRICE,M100.MCR_SS, M100.BEP_MAT_COST, M100.BEP_PROC_COST, M100.BEP_TOTAL_LOSS, M100.BEP_PROFIT_VALUE, M100.BEP_PRICE, M100.BEP_TARGET_PRICE, M100.PRODUCT_FINAL_PRICE FROM (SELECT DISTINCT G_CODE FROM ZTB_BOM2) AS AA JOIN M100 ON (AA.G_CODE = M100.G_CODE) JOIN M110 ON (M110.CUST_CD = M100.CUST_CD) WHERE (AA.G_CODE LIKE '%{keyword}%' OR G_NAME LIKE '%{keyword}%') "; 
            if(tinhgiachua == "chuatinhgia")
            {
                strQuery += "AND M100.PRODUCT_CMSPRICE is null";
            }


            result = config.GetData(strQuery);
            return result;

        }
        public DataTable traYCSXCODEKD(string keyword)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = $"SELECT P400.PROD_REQUEST_DATE, P400.PROD_REQUEST_NO,  M110.CUST_NAME_KD, P400.G_CODE, M100.G_NAME, P400.PROD_REQUEST_QTY, M010.EMPL_NAME, P400.EMPL_NO, P400.CUST_CD FROM P400 LEFT JOIN M100 ON (P400.G_CODE = M100.G_CODE) LEFT JOIN M010 ON (P400.EMPL_NO = M010.EMPL_NO) LEFT JOIN M110 ON (P400.CUST_CD = M110.CUST_CD) WHERE M100.G_NAME LIKE '%{keyword}%' ORDER BY P400.INS_DATE DESC";
            result = config.GetData(strQuery);
            return result;

        }


        public DataTable traYCSX(string codename,string fromdate, string todate)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "";
            if ((codename == ""))
            {
                //strQuery = "Select CTR_CD,PROD_REQUEST_DATE,PROD_REQUEST_NO,CODE_50,CODE_03,CODE_55,G_CODE,RIV_NO,PROD_REQUEST_QTY,CUST_CD,EMPL_NO,REMK,INS_EMPL,UPD_EMPL,DELIVERY_DT FROM P400";
                strQuery = "Select M100.G_NAME,P400.CTR_CD,P400.PROD_REQUEST_DATE,P400.PROD_REQUEST_NO,P400.CODE_50,P400.CODE_03,P400.CODE_55,P400.G_CODE,P400.RIV_NO,P400.PROD_REQUEST_QTY,P400.CUST_CD,P400.EMPL_NO,P400.REMK,P400.INS_EMPL,P400.UPD_EMPL,P400.DELIVERY_DT, CASE WHEN CODE_50 = '01' THEN 'GC' WHEN CODE_50 = '02' THEN 'SK' WHEN CODE_50 = '03' THEN 'KD' WHEN CODE_50 = '04' THEN 'VN' WHEN CODE_50 = '05' THEN 'Sample' WHEN CODE_50 = '06' THEN 'Vai bac 4' WHEN CODE_50 = '07' THEN 'ETC' END AS PHAN_LOAI_NHAP_KHAU, CASE WHEN CODE_55 = '01' THEN 'thong thuong' WHEN CODE_55 = '02' THEN 'sdi' WHEN CODE_55 = '03' THEN 'etc' WHEN CODE_55 = '04' THEN 'Sample' END AS PHAN_LOAI_SAN_XUAT  FROM P400 LEFT JOIN M100 ON P400.G_CODE = M100.G_CODE WHERE (PROD_REQUEST_DATE>=" + fromdate + " AND PROD_REQUEST_DATE<=" + todate + ")";
            }
            else
            {
                strQuery = "Select M100.G_NAME,P400.CTR_CD,P400.PROD_REQUEST_DATE,P400.PROD_REQUEST_NO,P400.CODE_50,P400.CODE_03,P400.CODE_55,P400.G_CODE,P400.RIV_NO,P400.PROD_REQUEST_QTY,P400.CUST_CD,P400.EMPL_NO,P400.REMK,P400.INS_EMPL,P400.UPD_EMPL,P400.DELIVERY_DT, CASE WHEN CODE_50 = '01' THEN 'GC' WHEN CODE_50 = '02' THEN 'SK' WHEN CODE_50 = '03' THEN 'KD' WHEN CODE_50 = '04' THEN 'VN' WHEN CODE_50 = '05' THEN 'Sample' WHEN CODE_50 = '06' THEN 'Vai bac 4' WHEN CODE_50 = '07' THEN 'ETC' END AS PHAN_LOAI_NHAP_KHAU, CASE WHEN CODE_55 = '01' THEN 'thong thuong' WHEN CODE_55 = '02' THEN 'sdi' WHEN CODE_55 = '03' THEN 'etc' WHEN CODE_55 = '04' THEN 'Sample' END AS PHAN_LOAI_SAN_XUAT  FROM P400 LEFT JOIN M100 ON P400.G_CODE = M100.G_CODE WHERE (G_NAME LIKE '%" + codename + "%' AND PROD_REQUEST_DATE>=" + fromdate + " AND PROD_REQUEST_DATE<="+ todate+")";
            }
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable traYCSXPIC(string picname, string codename, string fromdate, string todate)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "";
            if ((picname == ""))
            {
                //strQuery = "Select CTR_CD,PROD_REQUEST_DATE,PROD_REQUEST_NO,CODE_50,CODE_03,CODE_55,G_CODE,RIV_NO,PROD_REQUEST_QTY,CUST_CD,EMPL_NO,REMK,INS_EMPL,UPD_EMPL,DELIVERY_DT FROM P400";
                strQuery = "Select M100.G_NAME,P400.CTR_CD,P400.PROD_REQUEST_DATE,P400.PROD_REQUEST_NO,P400.CODE_50,P400.CODE_03,P400.CODE_55,P400.G_CODE,P400.RIV_NO,P400.PROD_REQUEST_QTY,P400.CUST_CD,P400.EMPL_NO,P400.REMK,P400.INS_EMPL,P400.UPD_EMPL,P400.DELIVERY_DT,CASE WHEN CODE_50 = '01' THEN 'GC' WHEN CODE_50 = '02' THEN 'SK' WHEN CODE_50 = '03' THEN 'KD' WHEN CODE_50 = '04' THEN 'VN' WHEN CODE_50 = '05' THEN 'Sample' WHEN CODE_50 = '06' THEN 'Vai bac 4' WHEN CODE_50 = '07' THEN 'ETC' END AS PHAN_LOAI_NHAP_KHAU, CASE WHEN CODE_55 = '01' THEN 'thong thuong' WHEN CODE_55 = '02' THEN 'sdi' WHEN CODE_55 = '03' THEN 'etc' WHEN CODE_55 = '04' THEN 'Sample' END AS PHAN_LOAI_SAN_XUAT  FROM P400 LEFT JOIN M100 ON P400.G_CODE = M100.G_CODE WHERE (PROD_REQUEST_DATE>=" + fromdate + " AND PROD_REQUEST_DATE<=" + todate + ")";
            }
            else
            {
                strQuery = "Select M100.G_NAME,P400.CTR_CD,P400.PROD_REQUEST_DATE,P400.PROD_REQUEST_NO,P400.CODE_50,P400.CODE_03,P400.CODE_55,P400.G_CODE,P400.RIV_NO,P400.PROD_REQUEST_QTY,P400.CUST_CD,P400.EMPL_NO,P400.REMK,P400.INS_EMPL,P400.UPD_EMPL,P400.DELIVERY_DT, CASE WHEN CODE_50 = '01' THEN 'GC' WHEN CODE_50 = '02' THEN 'SK' WHEN CODE_50 = '03' THEN 'KD' WHEN CODE_50 = '04' THEN 'VN' WHEN CODE_50 = '05' THEN 'Sample' WHEN CODE_50 = '06' THEN 'Vai bac 4' WHEN CODE_50 = '07' THEN 'ETC' END AS PHAN_LOAI_NHAP_KHAU, CASE WHEN CODE_55 = '01' THEN 'thong thuong' WHEN CODE_55 = '02' THEN 'sdi' WHEN CODE_55 = '03' THEN 'etc' WHEN CODE_55 = '04' THEN 'Sample' END AS PHAN_LOAI_SAN_XUAT  FROM P400 LEFT JOIN M100 ON P400.G_CODE = M100.G_CODE WHERE (G_NAME LIKE  '%" + codename + "%' AND EMPL_NO=  '" + picname + "' AND PROD_REQUEST_DATE>=" + fromdate + " AND PROD_REQUEST_DATE<=" + todate + ")";
            }
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable traLastYCSXPIC(string picname, string codename, string fromdate, string todate)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "";
            if ((picname == ""))
            {
                //strQuery = "Select CTR_CD,PROD_REQUEST_DATE,PROD_REQUEST_NO,CODE_50,CODE_03,CODE_55,G_CODE,RIV_NO,PROD_REQUEST_QTY,CUST_CD,EMPL_NO,REMK,INS_EMPL,UPD_EMPL,DELIVERY_DT FROM P400";
                strQuery = "Select M100.G_NAME,P400.CTR_CD,TOP 1 P400.PROD_REQUEST_DATE,P400.PROD_REQUEST_NO,P400.CODE_50,P400.CODE_03,P400.CODE_55,P400.G_CODE,P400.RIV_NO,P400.PROD_REQUEST_QTY,P400.CUST_CD,P400.EMPL_NO,P400.REMK,P400.INS_EMPL,P400.UPD_EMPL,P400.DELIVERY_DT,CASE WHEN CODE_50 = '01' THEN 'GC' WHEN CODE_50 = '02' THEN 'SK' WHEN CODE_50 = '03' THEN 'KD' WHEN CODE_50 = '04' THEN 'VN' WHEN CODE_50 = '05' THEN 'Sample' WHEN CODE_50 = '06' THEN 'Vai bac 4' WHEN CODE_50 = '07' THEN 'ETC' END AS PHAN_LOAI_NHAP_KHAU, CASE WHEN CODE_55 = '01' THEN 'thong thuong' WHEN CODE_55 = '02' THEN 'sdi' WHEN CODE_55 = '03' THEN 'etc' WHEN CODE_55 = '04' THEN 'Sample' END AS PHAN_LOAI_SAN_XUAT  FROM P400 LEFT JOIN M100 ON P400.G_CODE = M100.G_CODE WHERE (PROD_REQUEST_DATE>=" + fromdate + " AND PROD_REQUEST_DATE<=" + todate + ")";
            }
            else
            {
                strQuery = "Select M100.G_NAME,P400.CTR_CD,TOP 1 P400.PROD_REQUEST_DATE,P400.PROD_REQUEST_NO,P400.CODE_50,P400.CODE_03,P400.CODE_55,P400.G_CODE,P400.RIV_NO,P400.PROD_REQUEST_QTY,P400.CUST_CD,P400.EMPL_NO,P400.REMK,P400.INS_EMPL,P400.UPD_EMPL,P400.DELIVERY_DT, CASE WHEN CODE_50 = '01' THEN 'GC' WHEN CODE_50 = '02' THEN 'SK' WHEN CODE_50 = '03' THEN 'KD' WHEN CODE_50 = '04' THEN 'VN' WHEN CODE_50 = '05' THEN 'Sample' WHEN CODE_50 = '06' THEN 'Vai bac 4' WHEN CODE_50 = '07' THEN 'ETC' END AS PHAN_LOAI_NHAP_KHAU, CASE WHEN CODE_55 = '01' THEN 'thong thuong' WHEN CODE_55 = '02' THEN 'sdi' WHEN CODE_55 = '03' THEN 'etc' WHEN CODE_55 = '04' THEN 'Sample' END AS PHAN_LOAI_SAN_XUAT  FROM P400 LEFT JOIN M100 ON P400.G_CODE = M100.G_CODE WHERE (G_NAME LIKE  '%" + codename + "%' AND EMPL_NO=  '" + picname + "' AND PROD_REQUEST_DATE>=" + fromdate + " AND PROD_REQUEST_DATE<=" + todate + ")";
            }
            result = config.GetData(strQuery);
            return result;

        }





        public DataTable checkDuplicateYCSX(string ycsxno)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();           
            string strQuery = "SELECT PROD_REQUEST_NO FROM P400 WHERE PROD_REQUEST_NO=" + "'" + ycsxno + "'";
            result = config.GetData(strQuery);
            return result;

        }
        public DataTable getLastYCSXNo()
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = "SELECT TOP 1 PROD_REQUEST_NO FROM P400 ORDER BY INS_DATE DESC";
            result = config.GetData(strQuery);
            return result;

        }
        public DataTable getLastOutNo(string keyword)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = "SELECT MAX(OUT_NO) AS OUT_NO FROM O300 WHERE OUT_DATE='"+ keyword + "'";
            result = config.GetData(strQuery);
            return result;
        }
        public DataTable getLastOutSEQO301(string OUTDATE, string OUT_NO)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = $"SELECT TOP 1 OUT_SEQ FROM O301 WHERE OUT_DATE='" + OUTDATE + $"' AND OUT_NO='{OUT_NO}'";
            result = config.GetData(strQuery);
            return result;
        }

        public DataTable DeletePO(string PO_ID)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "DELETE FROM ZTBPOTable WHERE PO_ID=" + "'" + PO_ID + "'";
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable DeleteFCST(string FCST_ID)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "DELETE FROM ZTBFCSTTB WHERE FCST_ID=" + "'" + FCST_ID + "'";
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable DeleteInvoice(string DELIVERY_ID)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "DELETE FROM ZTBDelivery WHERE DELIVERY_ID=" + "'" + DELIVERY_ID + "'";
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable DeletePlan(string PLAN_ID)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "DELETE FROM ZTBPLANTB WHERE PLAN_ID=" + "'" + PLAN_ID + "'";
            result = config.GetData(strQuery);
            return result;

        }


        public DataTable DeleteYCSX(string ycsxno)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "DELETE FROM P400 WHERE PROD_REQUEST_NO=" + "'" + ycsxno + "'";
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable PheDuyetYCSX(string ycsxno)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = "UPDATE P400 SET PDUYET = 1 WHERE PROD_REQUEST_NO=" + "'" + ycsxno + "'";
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable checkYCSXO300(string ycsxno)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = $"SELECT TOP 1 * FROM O300 WHERE PROD_REQUEST_NO = '{ycsxno}'";
            result = config.GetData(strQuery);
            return result;
        }
        public DataTable checkvaokiem(string g_code)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = $"SELECT TOP 1 * FROM ZTBINSPECTNGTB WHERE G_CODE='{g_code}'";
            result = config.GetData(strQuery);
            return result;
        }

        public DataTable InsertO300(string insertValue)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = $"INSERT INTO O300 (CTR_CD,OUT_DATE, OUT_NO, CODE_03, CODE_50, CODE_52, PROD_REQUEST_DATE, PROD_REQUEST_NO, USE_YN, INS_DATE, INS_EMPL, FACTORY) VALUES {insertValue}";
            //MessageBox.Show(strQuery);
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable InsertO301(string insertValue)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = $"INSERT INTO O301 (CTR_CD,OUT_DATE, OUT_NO, OUT_SEQ, CODE_03, M_CODE, OUT_PRE_QTY, USE_YN, INS_DATE, INS_EMPL) VALUES {insertValue} ";
            //MessageBox.Show(strQuery);
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable InsertBOMAmazone(string insertValue)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = $"INSERT INTO BOM_AMAZONE (CTR_CD,G_CODE,G_CODE_MAU,DOITUONG_NO,GIATRI,REMARK,INS_DATE,INS_EMPL,UPD_DATE,UPD_EMPL) VALUES {insertValue} ";
            //MessageBox.Show(strQuery);
            result = config.GetData(strQuery);
            return result;

        }
        public DataTable InsertDESIGNAmazone(string insertValue)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = $"INSERT INTO DESIGN_AMAZONE (CTR_CD,G_CODE_MAU,DOITUONG_NO,DOITUONG_NAME,PHANLOAI_DT,CAVITY_PRINT,FONT_NAME,POS_X,POS_Y,SIZE_W,SIZE_H,ROTATE,REMARK,INS_DATE,INS_EMPL,UPD_DATE,UPD_EMPL,FONT_SIZE,FONT_STYLE,GIATRI) VALUES {insertValue} ";
            //MessageBox.Show(strQuery);
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable UpdateDESIGNAmazone(string updateValue)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = $"UPDATE DESIGN_AMAZONE {updateValue}";
            //MessageBox.Show(strQuery);
            result = config.GetData(strQuery);
            return result;

        }
        public DataTable DeleteDESIGNAmazone(string G_CODE_MAU)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = $"DELETE FROM DESIGN_AMAZONE WHERE G_CODE_MAU='{G_CODE_MAU}'";
            //MessageBox.Show(strQuery);
            result = config.GetData(strQuery);
            return result;

        }

        public DataTable UpdateBOMAmazone(string updateValue)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            //string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            string strQuery = $"UPDATE BOM_AMAZONE SET " + updateValue;
            //MessageBox.Show(strQuery);
            result = config.GetData(strQuery);
            return result;

        }




        // hàm getdata  product CMS
        public DataTable GetDataProductCMS(string item)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = "SELECT G_Code,G_Name FROM M100 where G_Code='" + item + "'";
            result = config.GetData(strQuery);
            return result;

        }


        // hàm getdata
        public DataTable GetData()
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = "SELECT ID_Product , Product_MaVach  ,Product_TenCode ,Product_TenCode_Full,PackingBox FROM tbl_Product ORDER BY ID_Product DESC";
            result = config.GetData(strQuery);
            return result;

        }

        // hàm getdata stock các sản phẩm
        public DataTable GetDataStock()
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = "SELECT  Product_MaVach ,Product_TenCode,QtyTon FROM tbl_Product ORDER BY Product_TenCode";
            result = config.GetData(strQuery);
            return result;

        }

        // hàm getdata stock các sản phẩm
        public DataTable GetDataStockAll()
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = "SELECT  Product_MaVach ,Product_TenCode,QtyTon FROM tbl_Product where QtyTon>0 ORDER BY Product_TenCode";
            result = config.GetData(strQuery);
            return result;

        }

        // hàm getdata stock các sản phẩm
        public DataTable GetDataSearchStock(string item)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = "SELECT  Product_MaVach ,Product_TenCode,QtyTon FROM tbl_Product where Product_MaVach LIKE '%" + item + "%' or Product_TenCode LIKE '%" + item + "%'  ORDER BY Product_TenCode";
            result = config.GetData(strQuery);
            return result;

        }

        // hàm tìm kiếm
        public DataTable SearchData(string item)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = "SELECT ID_Product , Product_MaVach  ,Product_TenCode ,Product_TenCode_Full,PackingBox FROM tbl_Product where Product_MaVach LIKE '%" + item + "%' or Product_TenCode LIKE '%" + item + "%' ORDER BY ID_Product DESC";
            result = config.GetData(strQuery);
            return result;

        }

        // hàm tìm kiếm
        public DataTable SearchDataLot(string item)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = "SELECT ID_Product, Product_MaVach  ,Product_TenCode ,PackingBox FROM tbl_Product where Product_MaVach LIKE '%" + item + "%' or Product_TenCode LIKE '%" + item + "%' ORDER BY Product_MaVach DESC";
            result = config.GetData(strQuery);
            return result;

        }


        // hàm hiển thị sản phẩn theo id
        public DataTable LoadDataByID_Product(string ID_Product)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = "SELECT * FROM tbl_Product where ID_Product =" + ID_Product;
            result = config.GetData(strQuery);
            return result;

        }

        // hàm hiển thị sản phẩn theo Product_MaVach
        public DataTable LoadDataByProduct_MaVach(string Product_MaVach)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = "SELECT * FROM tbl_Product where Product_MaVach =N'" + Product_MaVach + "'";
            result = config.GetData(strQuery);
            return result;

        }


        //check id
        public bool CheckID(int ID_Product)
        {
            //bool result = true;
            DataConfig config = new DataConfig();
            string strQuery = "select * from tbl_Product where ID_Product='" + ID_Product + "'";
            DataTable dt = new DataTable();
            dt = config.GetData(strQuery);
            if (dt.Rows.Count > 0)
            {
                return true;
            }
            return false;
        }

        //check mã vạch
        public bool CheckMaVach(string _mavach)
        {
            //bool result = true;
            DataConfig config = new DataConfig();
            string strQuery = "select * from tbl_Product where Product_MaVach ='" + _mavach + "'";
            DataTable dt = new DataTable();
            dt = config.GetData(strQuery);
            if (dt.Rows.Count > 0)
            {
                return true;
            }
            return false;
        }

        // hàm hiển thị tìm kiếm tên sản phẩm ở cms
        public DataTable SearchByProduct_TenSP(string tensp)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = "SELECT G_Code,G_Name,USE_YN FROM M100 where G_Name LIKE '%" + tensp + "%' order by G_Name ";
            result = config.GetData(strQuery);
            return result;

        }

        // hàm hiển thị tìm kiếm mã vạch sản phẩm ở cms
        public DataTable SearchByProduct_MVSP(string tensp)
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = "SELECT G_Code,G_Name,USE_YN FROM M100 where G_Code LIKE '%" + tensp + "%' order by G_Name ";
            result = config.GetData(strQuery);
            return result;

        }
        // hàm hiển thị tìm kiếm tên sản phẩm ở cms
        public DataTable LoadDataAllCMS()
        {
            DataTable result = new DataTable();
            DataConfig config = new DataConfig();
            string strQuery = "SELECT G_Code,G_Name,USE_YN FROM M100 order by G_Name  ";
            result = config.GetData(strQuery);
            return result;

        }

    }



    public class EncodeText
    {
        static string key = "nanioshitoruomae";
        public static string Encrypt(string toEncrypt)
        {
            bool useHashing = true;
            byte[] keyArray;
            byte[] toEncryptArray = UTF8Encoding.UTF8.GetBytes(toEncrypt);

            if (useHashing)
            {
                MD5CryptoServiceProvider hashmd5 = new MD5CryptoServiceProvider();
                keyArray = hashmd5.ComputeHash(UTF8Encoding.UTF8.GetBytes(key));
            }
            else
                keyArray = UTF8Encoding.UTF8.GetBytes(key);

            TripleDESCryptoServiceProvider tdes = new TripleDESCryptoServiceProvider();
            tdes.Key = keyArray;
            tdes.Mode = CipherMode.ECB;
            tdes.Padding = PaddingMode.PKCS7;

            ICryptoTransform cTransform = tdes.CreateEncryptor();
            byte[] resultArray = cTransform.TransformFinalBlock(toEncryptArray, 0, toEncryptArray.Length);

            return Convert.ToBase64String(resultArray, 0, resultArray.Length);
        }
        public static string Decrypt(string toDecrypt)
        {
            bool useHashing = true;
            byte[] keyArray;
            byte[] toEncryptArray = Convert.FromBase64String(toDecrypt);

            if (useHashing)
            {
                MD5CryptoServiceProvider hashmd5 = new MD5CryptoServiceProvider();
                keyArray = hashmd5.ComputeHash(UTF8Encoding.UTF8.GetBytes(key));
            }
            else
                keyArray = UTF8Encoding.UTF8.GetBytes(key);

            TripleDESCryptoServiceProvider tdes = new TripleDESCryptoServiceProvider();
            tdes.Key = keyArray;
            tdes.Mode = CipherMode.ECB;
            tdes.Padding = PaddingMode.PKCS7;

            ICryptoTransform cTransform = tdes.CreateDecryptor();
            byte[] resultArray = cTransform.TransformFinalBlock(toEncryptArray, 0, toEncryptArray.Length);
            return UTF8Encoding.UTF8.GetString(resultArray);
        }
    }



  


}
