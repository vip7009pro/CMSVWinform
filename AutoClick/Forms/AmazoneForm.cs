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
using System.Windows.Markup;

namespace AutoClick
{
    public partial class AmazoneForm : Form
    {
        public string Login_ID = "";
        public string PROD_REQUEST_NO = "";

        public int tra_data_flag = 0;
        public int check_data_flag = 0;
        public int upload_data_flag = 0;
        public int check_flag = 0;
        public DataTable dtgv1_data = new DataTable();

        public AmazoneForm()
        {
            InitializeComponent();
        }

        private void gradientPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void AmazoneForm_DragDrop(object sender, DragEventArgs e)
        {
           

        }

        public void handleDataAmazone(DoWorkEventArgs e)
        {
            if (dataGridView1.DataSource != null)
            {
                if (textBox4.Text != "" && textBox6.Text != "")
                {
                    int cavity_print = int.Parse(label4.Text);
                    string prod_request_no = textBox4.Text;
                    string g_code = label2.Text;
                    string no_in = textBox6.Text;
                    string data1 = "", data2 = "", data3 = "", data4 = "";
                    string status = "";
                    string inlai_count = "";
                    string remark = "";
                    string empl_no = "NHU1903";
                    DataTable dt = new DataTable();
                    dt.Columns.Add("G_CODE", typeof(string));
                    dt.Columns.Add("PROD_REQUEST_NO", typeof(string));
                    dt.Columns.Add("NO_IN", typeof(string));
                    dt.Columns.Add("ROW_NO", typeof(string));
                    dt.Columns.Add("DATA1", typeof(string));
                    dt.Columns.Add("DATA2", typeof(string));
                    dt.Columns.Add("DATA3", typeof(string));
                    dt.Columns.Add("DATA4", typeof(string));
                    dt.Columns.Add("STATUS", typeof(string));
                    dt.Columns.Add("INLAI_COUNT", typeof(string));
                    dt.Columns.Add("REMARK", typeof(string));


                    if (dataGridView1.Rows.Count % cavity_print == 0)
                    {
                        for (int i = 1; i <= dataGridView1.Rows.Count / cavity_print; i++)
                        {
                            DataRow dr1 = dt.NewRow();
                            dr1["G_CODE"] = g_code;
                            dr1["PROD_REQUEST_NO"] = prod_request_no;
                            dr1["NO_IN"] = no_in;
                            dr1["ROW_NO"] = i.ToString();
                            if (cavity_print == 1)
                            {
                                dr1["DATA1"] = dataGridView1.Rows[i].Cells["DATA"].Value.ToString();
                            }
                            else if (cavity_print == 2)
                            {
                                dr1["DATA1"] = dataGridView1.Rows[i * cavity_print - 1 - 1].Cells["DATA"].Value.ToString();
                                dr1["DATA2"] = dataGridView1.Rows[i * cavity_print - 1].Cells["DATA"].Value.ToString();
                            }
                            else if (cavity_print == 3)
                            {
                                dr1["DATA1"] = dataGridView1.Rows[i * cavity_print - 2 - 1].Cells["DATA"].Value.ToString();
                                dr1["DATA2"] = dataGridView1.Rows[i * cavity_print - 1 - 1].Cells["DATA"].Value.ToString();
                                dr1["DATA3"] = dataGridView1.Rows[i * cavity_print - 1].Cells["DATA"].Value.ToString();
                            }
                            else if (cavity_print == 4)
                            {
                                dr1["DATA1"] = dataGridView1.Rows[i * cavity_print - 3 - 1].Cells["DATA"].Value.ToString();
                                dr1["DATA2"] = dataGridView1.Rows[i * cavity_print - 2 - 1].Cells["DATA"].Value.ToString();
                                dr1["DATA3"] = dataGridView1.Rows[i * cavity_print - 1 - 1].Cells["DATA"].Value.ToString();
                                dr1["DATA4"] = dataGridView1.Rows[i * cavity_print - 1].Cells["DATA"].Value.ToString();
                            }
                            dr1["INLAI_COUNT"] = "0";
                            dr1["REMARK"] = "0";
                            dt.Rows.Add(dr1);
                        }
                        dataGridView1.DataSource = dt;
                        new Form1().setRowNumber(dataGridView1);
                        
                        ProductBLL pro = new ProductBLL();

                        progressBar1.Minimum = 0; //Đặt giá trị nhỏ nhất cho ProgressBar
                        progressBar1.Maximum = dataGridView1.Rows.Count; //Đặt giá trị lớn nhất cho ProgressBar
                        int startprogress = 0;

                        //insert vao database
                        for (int r = 0; r < dataGridView1.Rows.Count; r++)
                        {
                            DataGridViewRow row = dataGridView1.Rows[r];
                            if (!row.IsNewRow)
                            {
                                string G_CODE = row.Cells["G_CODE"].Value.ToString();
                                string NO_IN = row.Cells["NO_IN"].Value.ToString();
                                string PROD_REQUEST_NO = row.Cells["PROD_REQUEST_NO"].Value.ToString();
                                string ROW_NO = row.Cells["ROW_NO"].Value.ToString();
                                string DATA1 = row.Cells["DATA1"].Value.ToString();
                                string DATA2 = row.Cells["DATA2"].Value.ToString();
                                string DATA3 = row.Cells["DATA3"].Value.ToString();
                                string DATA4 = row.Cells["DATA4"].Value.ToString();
                                string STATUS = row.Cells["STATUS"].Value.ToString();
                                string INLAI_COUNT = row.Cells["INLAI_COUNT"].Value.ToString();
                                string REMARK = row.Cells["REMARK"].Value.ToString();
                                string EMPL_NO = Login_ID;
                                string insertValue = $"('002','{G_CODE}','{PROD_REQUEST_NO}','{NO_IN}','{ROW_NO}','{DATA1}','{DATA2}','{DATA3}','{DATA4}','{STATUS}','{INLAI_COUNT}', '0912',GETDATE(),'{Login_ID}', GETDATE(),'{Login_ID}')";
                                try
                                {
                                    pro.insertAMAZONEDATA(insertValue);
                                    //MessageBox.Show(insertValue);
                                    backgroundWorker1.ReportProgress(r);
                                    dataGridView1.CurrentCell = dataGridView1.Rows[r].Cells[0];
                                }
                                catch(Exception  ex)
                                {
                                    MessageBox.Show("Lỗi: " + ex.ToString());
                                }
                                
                                //label5.Text = "Progress: " + startprogress + "/" + dataGridView1.Rows.Count;
                                //progressBar1.Value = startprogress;
                            }
                        }
                        progressBar1.Value = 0;
                        MessageBox.Show("Upload data amazone thành công !");
                        check_flag = 0;                        
                    }
                    else
                    {
                        MessageBox.Show("Data lẻ so với cavity in, kiểm tra lại data");
                    }

                }
                else
                {
                    MessageBox.Show("Không được để trống ô nào !");
                }
            }
            else
            {
                MessageBox.Show("Kéo file excel vào đã!");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
          updataBT();            
        }

        private void dataGridView1_DragDrop(object sender, DragEventArgs e)
        {
            if(textBox6.Text =="")
            {
                MessageBox.Show("Phải nhập ID công việc trước khi kéo file");
            }
            else if(textBox4.Text =="")
            {
                MessageBox.Show("Phải nhập YCSX trước khi kéo file");
            }
            else
            {
                DataTable dt = new DataTable();
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop, false);
                foreach (string file in files)
                {
                    //MessageBox.Show(file);
                    //MessageBox.Show(textBox6.Text);
                    string G_CODE = label2.Text;
                    string model_name = "";
                    ProductBLL pro = new ProductBLL();
                    DataTable dt1 = new DataTable();                    
                    dt1= pro.checkModelNameAmazone(G_CODE);
                    if(dt1.Rows.Count <= 0)
                    {
                        //MessageBox.Show("Rnd chưa nhập model của code này lên hệ thống, ko thể so sánh model");
                    } 
                    else
                    {
                        model_name = dt1.Rows[0]["PROD_MODEL"].ToString();
                    }


                    if (dt1.Rows.Count <= 0)
                    {
                        MessageBox.Show("Rnd chưa nhập model của code này lên hệ thống, ko thể so sánh model");
                    }
                    else if(!file.Contains(model_name))
                    {
                        MessageBox.Show("Nghi ngờ kéo nhầm file so với ID công việc : " + textBox6.Text + "| Phát hiện sai model ! ");
                    }
                    else if (!file.Contains(textBox6.Text))
                    {
                        MessageBox.Show("Nghi ngờ kéo nhầm file so với ID công việc : " + textBox6.Text);
                    }
                    else
                    {
                        try
                        {
                            if (file != "")
                            {
                                dt = ExcelFactory.readFromExcelFileToAmazoneTable(file);
                                this.dataGridView1.DataSource = null;
                                this.dataGridView1.Columns.Clear();
                                dataGridView1.DataSource = dt;

                                bool error3 = false;
                                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                                {
                                    for (int j = 0; j < dataGridView1.Rows.Count; j++)
                                    {
                                        if (i != j)
                                        {

                                            if (dataGridView1.Rows[i].Cells["DATA"].Value.ToString() == dataGridView1.Rows[j].Cells["DATA"].Value.ToString())
                                            {
                                                error3 = true;
                                            }

                                        }

                                    }
                                }

                                if(!error3)
                                {
                                    dataGridView1.DataSource = dt;
                                }
                                else
                                {
                                    dataGridView1.DataSource = null;
                                    MessageBox.Show("File kéo vào có dòng trùng");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.ToString());
                        }
                        MessageBox.Show("Import file hoàn thành!, có " + (dataGridView1.Rows.Count) + " dòng");
                    }
                    
                }

            }          
           
        }

        private void AmazoneForm_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        private void dataGridView1_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        private void AmazoneForm_Load(object sender, EventArgs e)
        {
            if (!System.Windows.Forms.SystemInformation.TerminalServerSession)
            {
                Type dgvType = dataGridView1.GetType();
                PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                  BindingFlags.Instance | BindingFlags.NonPublic);
                pi.SetValue(dataGridView1, true, null);
            }
            
            textBox4.Text = PROD_REQUEST_NO;
            pictureBox1.Hide();
        }

        private void textBox4_Leave(object sender, EventArgs e)
        {
            try
            {
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();
                dt = pro.getYCSXInfo2(textBox4.Text);
                label1.Text = "CODE: " + dt.Rows[0]["G_NAME"].ToString();
                label2.Text =  dt.Rows[0]["G_CODE"].ToString();
                dt = pro.getcavity_print(label2.Text);
                label4.Text = dt.Rows[0]["CAVITY_PRINT"].ToString();
            }
            catch(Exception ex)
            {
                //MessageBox.Show("Lỗi: " + ex.ToString());
            }           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string RunningPath = AppDomain.CurrentDomain.BaseDirectory;
            string FileName = string.Format("{0}Resources\\darkgreen.jpg", Path.GetFullPath(Path.Combine(RunningPath, @"..\..\")));
            MessageBox.Show(FileName);
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            try
            {
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();
                dt = pro.getYCSXInfo2(textBox4.Text);
                label1.Text = "CODE: " + dt.Rows[0]["G_NAME"].ToString();
                label2.Text = dt.Rows[0]["G_CODE"].ToString();
                dt = pro.getcavity_print(label2.Text);
                label4.Text = dt.Rows[0]["CAVITY_PRINT"].ToString();
                //label4.Text = "2";
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Lỗi: " + ex.ToString());
            }
        }

        public void tradataBT()
        {
            tra_data_flag = 1;
            check_data_flag = 0;
            upload_data_flag = 0;
            if (!backgroundWorker1.IsBusy)
            {
                pictureBox1.Show();
                backgroundWorker1.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Đang trong tiến trình khác, thử lại sau");
            }

        }

        public void checkdataBT()
        {
            tra_data_flag = 0;
            check_data_flag = 1;
            upload_data_flag = 0;
            if (!backgroundWorker1.IsBusy)
            {
                pictureBox1.Show();
                backgroundWorker1.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Đang trong tiến trình khác, thử lại sau");
            }

        }
        public void updataBT()
        {
            tra_data_flag = 0;
            check_data_flag = 0;
            upload_data_flag = 1;
            if (!backgroundWorker1.IsBusy)
            {
                pictureBox1.Show();
                backgroundWorker1.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Đang trong tiến trình khác, thử lại sau");
            }

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            tradataBT();
            
        }

        public int checkDataAmazonSau (DoWorkEventArgs e)
        {
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();            
            
            dt = pro.checkAMZDuplicateCount();
            if(dt.Rows.Count > 0)
            {
                return 0; 
            }
            else
            {
                return 1;
            }


        }
        public void checkDataAmazone(DoWorkEventArgs e)
        {
            bool error3 = false;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Rows.Count; j++)
                {
                    if (i != j)
                    {

                        if (dataGridView1.Rows[i].Cells["DATA"].Value.ToString() == dataGridView1.Rows[j].Cells["DATA"].Value.ToString())
                        {
                            error3 = true;
                        }

                    }

                }
            }
            if(!error3)
            {
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();
                if (dataGridView1.Rows.Count > 0)
                {
                    bool error1 = false;
                    bool error2 = false;
                    

                    /*for (int r = 0; r < dataGridView1.Rows.Count; r++)
                    {
                        DataGridViewRow row = dataGridView1.Rows[r];
                        if (!row.IsNewRow)
                        {
                            string DATA = row.Cells["DATA"].Value.ToString();
                            dt = pro.checkDataAmazone(DATA);
                            dataGridView1.Rows[row.Index].Cells["DATA"].Style.BackColor = Color.LightGreen;
                            backgroundWorker1.ReportProgress(r);
                            dataGridView1.CurrentCell = dataGridView1.Rows[r].Cells[0];
                            if (dt.Rows.Count > 0)
                            {
                                error1 = true;
                                dataGridView1.Rows[row.Index].Cells["DATA"].Style.BackColor = Color.Red;
                            }
                        }
                    }
                    */

                    dt = pro.checkNO_IN_Amazone(textBox6.Text);
                    if (dt.Rows.Count > 0)
                    {
                        error2 = true;
                    }
                   
                    if (error2)
                    {
                        MessageBox.Show("ID công việc đã được sử dụng, không thể up trùng được");
                    }                   
                    if (!error2 && !error3)
                    {
                        MessageBox.Show("Check data thành công, bạn có thể upload data này");
                        check_flag = 1;
                    }

                }

            }
            else
            {
                MessageBox.Show("File kéo vào có dòng trùng, check lại file");
            }


        }
        private void button3_Click(object sender, EventArgs e)
        {
            checkdataBT();            
        }

        public void xuly(DoWorkEventArgs e)
        {
            if (tra_data_flag == 1)
            {
                ProductBLL pro = new ProductBLL();                
                if (textBox6.Text == "" && textBox4.Text == "")
                {
                    MessageBox.Show("Hãy nhập số yêu cầu, hoặc tên code, hoặc ID công việc");
                }
                else
                {
                    dtgv1_data = pro.checkDATAAmazone(textBox1.Text, textBox6.Text, textBox4.Text);                   
                }
            }
            else if(check_data_flag == 1)
            {               
                checkDataAmazone(e);                
            }
            else if(upload_data_flag == 1)
            {
                if (check_flag == 1)
                {
                    handleDataAmazone(e);
                    if(checkDataAmazonSau(e)==1)
                    {
                        MessageBox.Show("Upload data thành công");
                    }
                    else
                    {
                        MessageBox.Show("Data upload có vấn đề, hãy báo cáo");
                    }    
                }
                else
                {
                    MessageBox.Show("Check data trước khi bấm upload");
                }
                

               
            }

        }


        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            xuly(e);
            
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Maximum = dataGridView1.Rows.Count;
            if (tra_data_flag == 1)
            {
               
            }
            else if (check_data_flag == 1)
            {
                //progressBar1.Value = e.ProgressPercentage / dataGridView1.Rows.Count*100;
                progressBar1.Value = e.ProgressPercentage;
                label5.Text = "Progress: " + e.ProgressPercentage + "/" + dataGridView1.Rows.Count;
            }
            else if (upload_data_flag == 1)
            {
                //progressBar1.Value = e.ProgressPercentage / dataGridView1.Rows.Count*100;
                progressBar1.Value = e.ProgressPercentage;
                label5.Text = "Progress: " + e.ProgressPercentage + "/" + dataGridView1.Rows.Count;
            }            
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            pictureBox1.Hide();
            if (tra_data_flag == 1)
            {
                dataGridView1.DataSource = dtgv1_data;
                MessageBox.Show("Đã load: " + dtgv1_data.Rows.Count + " dòng");
            }
            else if (check_data_flag == 1)
            {
                
            }
            else if (upload_data_flag == 1)
            {
                
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dt = pro.checkAMZDuplicateCount();

            if (dt.Rows.Count > 0)
            {
                MessageBox.Show("Có dòng trùng");
                dataGridView1.DataSource = dt;
            }
            else
            {
                MessageBox.Show("Không có dòng trùng");
            }
        }
    }
}
