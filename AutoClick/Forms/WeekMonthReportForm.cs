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
    public partial class WeekMonthReportForm : Form
    {
        public WeekMonthReportForm()
        {
            InitializeComponent();
        }

        public void tradulieu()
        {
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            DataTable dttong = new DataTable();
            DataTable dtpic = new DataTable();
            DataTable dtdaily = new DataTable();

            dt = pro.report_CustomerDeliveryByType(STYMD2(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day), STYMD2(dateTimePicker2.Value.Year, dateTimePicker2.Value.Month, dateTimePicker2.Value.Day));
            dttong = pro.traInvoiceTotal(conditiongen());
            dtpic = pro.traInvoiceByPIC(STYMD2(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day), STYMD2(dateTimePicker2.Value.Year, dateTimePicker2.Value.Month, dateTimePicker2.Value.Day));
            dtdaily = pro.customerDaily(STYMD2(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day), STYMD2(dateTimePicker2.Value.Year, dateTimePicker2.Value.Month, dateTimePicker2.Value.Day));
            try
            {


                int dailycolumnum = dtdaily.Columns.Count;


                DataRow drpic = dtpic.NewRow();
                int colnumdtpic = dtpic.Columns.Count;
                int rownumdtpic = dtpic.Rows.Count;
                double sumdtpic = 0.0;
                for (int i = 1; i < colnumdtpic; i++)
                {
                    for (int j = 0; j < rownumdtpic; j++)
                    {
                        sumdtpic += double.Parse(dtpic.Rows[j][dtpic.Columns[i]].ToString());
                    }
                    drpic[i] = sumdtpic;
                    sumdtpic = 0;
                }
                drpic[0] = "TOTAL";
                dtpic.Rows.InsertAt(drpic, 0);


                DataRow drdt = dt.NewRow();
                int colnumdt = dt.Columns.Count;
                int rownumdt = dt.Rows.Count;
                double sumdt = 0.0;
                for (int i = 1; i < colnumdt; i++)
                {
                    for (int j = 0; j < rownumdt; j++)
                    {
                        sumdt += double.Parse(dt.Rows[j][dt.Columns[i]].ToString());
                    }
                    drdt[i] = sumdt;
                    sumdt = 0;
                }
                drdt[0] = "TOTAL";
                dt.Rows.InsertAt(drdt, 0);


                DataRow dr = dtdaily.NewRow();
                int colnum = dtdaily.Columns.Count;
                int rownum = dtdaily.Rows.Count;
                double sum = 0;
                for (int i = 1; i < colnum; i++)
                {
                    for (int j = 0; j < rownum; j++)
                    {
                        sum += double.Parse(dtdaily.Rows[j][dtdaily.Columns[i]].ToString());
                    }
                    dr[i] = sum;
                    sum = 0;
                }
                dr[0] = "TOTAL";
                dtdaily.Rows.InsertAt(dr, 0);

                /*

               int dailycolumnum = dtdaily.Columns.Count;
               if(dailycolumnum >=3)
               {


                   DateTime tempdate = dateTimePicker1.Value;
                   for (int i = 2; i < dailycolumnum; i++)
                   {
                       dtdaily.Columns[i].ColumnName = STYMD2(tempdate.Year, tempdate.Month, tempdate.Day);
                       tempdate = tempdate.AddDays(1);
                   }

               }
               */

                dataGridView1.DataSource = dt;
                dataGridView2.DataSource = dtpic;


                dataGridView3.Refresh();
                BindingSource BS = new BindingSource();
                BS.DataSource = dtdaily;
                dataGridView3.DataSource = BS;


                setRowNumber(dataGridView1);
                setRowNumber(dataGridView2);
                setRowNumber(dataGridView3);

                formatDataGridViewtraPO1(dataGridView1);
                formatDataGridViewtraInvoicePIC1(dataGridView2);
                if (dailycolumnum >= 3)
                {
                    formatDataGridViewtradailyClosing1(dataGridView3, dailycolumnum);
                }

                if (dt.Rows.Count > 0)
                {
                    textBox4.Text = dttong.Rows[0]["DELIVERED_QTY"].ToString();
                    textBox4.Text = string.Format("{0:#,##0 EA}", double.Parse(textBox4.Text));

                    textBox5.Text = dttong.Rows[0]["DELIVERED_AMOUNT"].ToString();
                    textBox5.Text = string.Format("{0:#,##0.00 $}", double.Parse(textBox5.Text));

                    textBox4.Font = new Font("Microsoft Sans Serif", 15.0f);
                    textBox5.Font = new Font("Microsoft Sans Serif", 15.0f);
                }
                else
                {

                }


            }
            catch(Exception ex)
            {
                MessageBox.Show("Không có data");
            }

        }

        public string conditiongen()
        {
            string query = "WHERE ";
            string fromdate, todate;
            fromdate = STYMD2(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day);
            todate = STYMD2(dateTimePicker2.Value.Year, dateTimePicker2.Value.Month, dateTimePicker2.Value.Day);
            //MessageBox.Show(fromdate);
            string ngaythang = "DELIVERY_DATE BETWEEN '" + fromdate + "' AND '" + todate + "' ";
            query = query + ngaythang;
            return query;
        }

        public void formatDataGridViewtradailyClosing1(DataGridView dataGridView1, int columnum)
        {
            

            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;            

            

            dataGridView1.Columns["DELIVERED_AMOUNT"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["DELIVERED_AMOUNT"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["DELIVERED_AMOUNT"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);
           
            for (int i = 1; i < columnum; i++)
            {
                dataGridView1.Columns[i].DefaultCellStyle.Format = "c";               
            }

            dataGridView1.Rows[0].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Rows[0].DefaultCellStyle.BackColor = Color.Green;
            dataGridView1.Rows[0].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);



        }


        public void formatDataGridViewtraInvoicePIC1(DataGridView dataGridView1)
        {
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;
            dataGridView1.Columns["DELIVERY_QTY"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["DELIVERED_AMOUNT"].DefaultCellStyle.Format = "c";
            dataGridView1.Columns["DELIVERY_QTY"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["DELIVERY_QTY"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["DELIVERY_QTY"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["DELIVERED_AMOUNT"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["DELIVERED_AMOUNT"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["DELIVERED_AMOUNT"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Rows[0].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Rows[0].DefaultCellStyle.BackColor = Color.Green;
            dataGridView1.Rows[0].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);


        }
        public void formatDataGridViewtraPO1(DataGridView dataGridView1)
        {
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;

            dataGridView1.Columns["DELIVERY_QTY"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["TSP_QTY"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["LABEL_QTY"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["UV_QTY"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["OLED_QTY"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["TAPE_QTY"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["RIBBON_QTY"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["SPT_QTY"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["OTHERS_QTY"].DefaultCellStyle.Format = "#,0";

            dataGridView1.Columns["DELIVERED_AMOUNT"].DefaultCellStyle.Format = "c";
            dataGridView1.Columns["TSP_AMOUNT"].DefaultCellStyle.Format = "c";
            dataGridView1.Columns["LABEL_AMOUNT"].DefaultCellStyle.Format = "c";
            dataGridView1.Columns["UV_AMOUNT"].DefaultCellStyle.Format = "c";
            dataGridView1.Columns["OLED_AMOUNT"].DefaultCellStyle.Format = "c";
            dataGridView1.Columns["TAPE_AMOUNT"].DefaultCellStyle.Format = "c";
            dataGridView1.Columns["RIBBON_AMOUNT"].DefaultCellStyle.Format = "c";
            dataGridView1.Columns["SPT_AMOUNT"].DefaultCellStyle.Format = "c";
            dataGridView1.Columns["OTHERS_AMOUNT"].DefaultCellStyle.Format = "c";           

            dataGridView1.Columns["DELIVERY_QTY"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["DELIVERY_QTY"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["DELIVERY_QTY"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["DELIVERED_AMOUNT"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["DELIVERED_AMOUNT"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["DELIVERED_AMOUNT"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["CUST_NAME_KD"].DefaultCellStyle.BackColor = Color.Aqua;

            dataGridView1.Rows[0].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Rows[0].DefaultCellStyle.BackColor = Color.Green;
            dataGridView1.Rows[0].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);


        }

        private void button1_Click(object sender, EventArgs e)
        {
            tradulieu();
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

        private void setRowNumber(DataGridView dgv)
        {
            foreach (DataGridViewRow row in dgv.Rows)
            {
                row.HeaderCell.Value = (row.Index + 1).ToString();
            }
        }


        private void WeekMonthReportForm_Load(object sender, EventArgs e)
        {
            //tradulieu();
            this.ContextMenuStrip = contextMenuStrip1;
            this.dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            this.dataGridView2.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            this.dataGridView3.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            if (!System.Windows.Forms.SystemInformation.TerminalServerSession)
            {
                Type dgvType = dataGridView1.GetType();
                PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                  BindingFlags.Instance | BindingFlags.NonPublic);
                pi.SetValue(dataGridView1, true, null);
            }
            if (!System.Windows.Forms.SystemInformation.TerminalServerSession)
            {
                Type dgvType = dataGridView1.GetType();
                PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                  BindingFlags.Instance | BindingFlags.NonPublic);
                pi.SetValue(dataGridView2, true, null);
            }
            if (!System.Windows.Forms.SystemInformation.TerminalServerSession)
            {
                Type dgvType = dataGridView1.GetType();
                PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                  BindingFlags.Instance | BindingFlags.NonPublic);
                pi.SetValue(dataGridView3, true, null);
            }
        }

        private void save1ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new Form1().CopyToClipboardWithHeaders(dataGridView1);
        }

        private void save2ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new Form1().CopyToClipboardWithHeaders(dataGridView2);
        }

        private void save3ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new Form1().CopyToClipboardWithHeaders(dataGridView3);
        }
    }
}
