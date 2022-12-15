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
    public partial class Chart1 : Form
    {
        public Chart1()
        {
            InitializeComponent();
        }

        private void Chart1_Load(object sender, EventArgs e)
        {           
           
                        
        }
        private void chart2_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            DataTable dttong = new DataTable();
            string fromdate = new Form1().STYMD2(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day);
            string todate = new Form1().STYMD2(dateTimePicker2.Value.Year, dateTimePicker2.Value.Month, dateTimePicker2.Value.Day);            
            dt = pro.traDelivery(fromdate, todate);

            chart2.DataSource = dt;
            chart2.DataBind();
            chart2.ChartAreas[0].AxisX.MajorGrid.LineDashStyle = System.Windows.Forms.DataVisualization.Charting.ChartDashStyle.DashDotDot;
            chart2.ChartAreas[0].AxisY.MajorGrid.LineDashStyle = System.Windows.Forms.DataVisualization.Charting.ChartDashStyle.DashDotDot;
            chart2.ChartAreas[0].AxisY2.MajorGrid.LineDashStyle = System.Windows.Forms.DataVisualization.Charting.ChartDashStyle.DashDotDot;
            chart2.ChartAreas[0].AxisY.LabelStyle.Format = "{0:0,}K $";
            chart2.ChartAreas[0].AxisY2.LabelStyle.Format = "{#,###,}K EA";

            chart4.DataSource = dt;
            chart4.DataBind();

            dt = pro.traDelivery_customer(fromdate, todate);
            chart3.DataSource = dt;    
            chart3.DataBind();



            dttong = pro.traPOTotal("");         
   
            textBox4.Text = dttong.Rows[0]["TOTAL_DELIVERED"].ToString();
            textBox4.Text = string.Format("{0:#,##0 EA}", double.Parse(textBox4.Text));
            textBox6.Text = dttong.Rows[0]["PO_BALANCE"].ToString();
            textBox6.Text = string.Format("{0:#,##0 EA}", double.Parse(textBox6.Text));
          
            textBox5.Text = dttong.Rows[0]["DELIVERED_AMOUNT"].ToString();
            textBox5.Text = string.Format("{0:#,##0.00 $}", double.Parse(textBox5.Text));
            textBox7.Text = dttong.Rows[0]["BALANCE_AMOUNT"].ToString();
            textBox7.Text = string.Format("{0:#,##0.00 $}", double.Parse(textBox7.Text));
                        
            textBox4.Font = new Font("Microsoft Sans Serif", 15.0f);
            textBox6.Font = new Font("Microsoft Sans Serif", 15.0f);
            
            textBox5.Font = new Font("Microsoft Sans Serif", 15.0f);
            textBox7.Font = new Font("Microsoft Sans Serif", 15.0f);   

        }
    }
}
