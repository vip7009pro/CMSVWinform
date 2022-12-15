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
    public partial class LICHSUSANXUAT : Form
    {
        public LICHSUSANXUAT()
        {
            InitializeComponent();
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

        public string tralichsuinputlieu_condition()
        {
            string condition = " WHERE ";

            string fromdate, todate;
            fromdate = STYMD(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day);
            todate = STYMD(dateTimePicker2.Value.Year, dateTimePicker2.Value.Month, dateTimePicker2.Value.Day);
            //MessageBox.Show(fromdate);
            string ngaythang = " P500.PROCESS_IN_DATE BETWEEN '" + fromdate + "' AND '" + todate + "' ";
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
                ycsxno = "AND P500.PROD_REQUEST_NO= '" + textBox4.Text + "' ";
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
            string cmscode = "";
            if (textBox2.Text != "")
            {
                cmscode = "AND M100.G_CODE='" + textBox2.Text + "'";
            }
            else
            {
                cmscode = "";
            }
            condition += ngaythang + code + ycsxno +  cmscode + picname;

            return condition;
        }
        private void button1_Click(object sender, EventArgs e)
        {          

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

        public void formatYCSXTable(DataGridView dataGridView1)
        {
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;

            

            dataGridView1.Columns["G_NAME"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.BackColor = Color.Black;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["PROD_REQUEST_NO"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["PROD_REQUEST_NO"].DefaultCellStyle.BackColor = Color.Black;
            dataGridView1.Columns["PROD_REQUEST_NO"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

        }

        private void LICHSUSANXUAT_Load(object sender, EventArgs e)
        {
            pictureBox1.Hide();          
            if (!System.Windows.Forms.SystemInformation.TerminalServerSession)
            {
                Type dgvType = dataGridView2.GetType();
                PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                  BindingFlags.Instance | BindingFlags.NonPublic);
                pi.SetValue(dataGridView2, true, null);
            }
        }
        public DataTable dt = new DataTable();
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            ProductBLL pro = new ProductBLL();           
            dt = pro.traLICHSUINPUTLIEU(tralichsuinputlieu_condition());           
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            pictureBox1.Hide();
            dataGridView2.Columns.Clear();
            dataGridView2.DataSource = dt;
            formatYCSXTable(dataGridView2);
            MessageBox.Show("Đã load : " + dt.Rows.Count + " dòng");
        }
    }
}
