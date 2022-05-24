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
using System.Reflection;

namespace AutoClick
{
    public partial class SOPForm : Form
    {
        public SOPForm()
        {
            InitializeComponent();
        }



        public void load_SOP()
        {
            this.dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dt = pro.report_SOP(STYMD2(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day));

            int dailycolumnum = dt.Columns.Count;
            if (dailycolumnum >= 7)
            {
                DateTime tempdate = dateTimePicker1.Value;
                for (int i = 7; i < dailycolumnum; i++)
                {
                    dt.Columns[i].ColumnName = STYMD2(tempdate.Year, tempdate.Month, tempdate.Day);
                    tempdate = tempdate.AddDays(1);
                }
            }

            DataRow dr = dt.NewRow();
            int colnum = dt.Columns.Count;
            int rownum = dt.Rows.Count;
            int sum = 0;
            for (int i = 6; i < colnum; i++)
            {
                for (int j = 0; j < rownum; j++)
                {
                    sum += int.Parse(dt.Rows[j][dt.Columns[i]].ToString());
                }
                dr[i] = sum;
                sum = 0;
            }
            dr[5] = "TOTAL";
            dt.Rows.InsertAt(dr, 0);
            MessageBox.Show("Người đã up PLAN: \n" + pro.report_SOP_uploadedPIC(STYMD2(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day)));

            dataGridView1.DataSource = dt;
            setRowNumber(dataGridView1);
            formatDataGridViewtraInvoicePIC1(dataGridView1, dt);
            if (!System.Windows.Forms.SystemInformation.TerminalServerSession)
            {
                Type dgvType = dataGridView1.GetType();
                PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                  BindingFlags.Instance | BindingFlags.NonPublic);
                pi.SetValue(dataGridView1, true, null);
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
                try
                {  
                    load_SOP();                    
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Xảy ra lỗi, chi tiết : \n" + ex.ToString());
                }
          
        }



        public static void FastAutoSizeColumns(DataGridView targetGrid)
        {
            // Cast out a DataTable from the target grid datasource.
            // We need to iterate through all the data in the grid and a DataTable supports enumeration.
            var gridTable = (DataTable)targetGrid.DataSource;

            // Create a graphics object from the target grid. Used for measuring text size.
            using (var gfx = targetGrid.CreateGraphics())
            {
                // Iterate through the columns.
                for (int i = 0; i < gridTable.Columns.Count; i++)
                {
                    // Leverage Linq enumerator to rapidly collect all the rows into a string array, making sure to exclude null values.
                    string[] colStringCollection = gridTable.AsEnumerable().Where(r => r.Field<object>(i) != null).Select(r => r.Field<object>(i).ToString()).ToArray();

                    // Sort the string array by string lengths.
                    colStringCollection = colStringCollection.OrderBy((x) => x.Length).ToArray();

                    // Get the last and longest string in the array.
                    string longestColString = colStringCollection.Last();

                    // Use the graphics object to measure the string size.
                    var colWidth = gfx.MeasureString(longestColString, targetGrid.Font);

                    // If the calculated width is larger than the column header width, set the new column width.
                    if (colWidth.Width > targetGrid.Columns[i].HeaderCell.Size.Width)
                    {
                        targetGrid.Columns[i].Width = (int)colWidth.Width;
                    }
                    else // Otherwise, set the column width to the header width.
                    {
                        targetGrid.Columns[i].Width = targetGrid.Columns[i].HeaderCell.Size.Width;
                    }
                }
            }
        }



        public void formatDataGridViewtraInvoicePIC1(DataGridView dataGridView1, DataTable dt)
        {
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;

            dataGridView1.Columns["PO_BALANCE"].DefaultCellStyle.Format = "#,0";

            int dailycolumnum = dt.Columns.Count;
            if (dailycolumnum >= 7)
            {
                DateTime tempdate = dateTimePicker1.Value;
                for (int i = 7; i < dailycolumnum; i++)
                {
                    dt.Columns[i].ColumnName = STYMD2(tempdate.Year, tempdate.Month, tempdate.Day);
                    dataGridView1.Columns[i].DefaultCellStyle.Format = "#,0";
                    tempdate = tempdate.AddDays(1);
                }
            }


            dataGridView1.Columns["PO_BALANCE"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["PO_BALANCE"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["PO_BALANCE"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Rows[0].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Rows[0].DefaultCellStyle.BackColor = Color.Green;
            dataGridView1.Rows[0].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);
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

        private void SOPForm_Load(object sender, EventArgs e)
        {

        }
        private void setRowNumber(DataGridView dgv)
        {
            foreach (DataGridViewRow row in dgv.Rows)
            {
                row.HeaderCell.Value = (row.Index + 1).ToString();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelFactory.writeToExcelFile(dataGridView1);
        }
    }
}
