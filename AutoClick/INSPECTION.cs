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
    public partial class INSPECTION : Form
    {
        public INSPECTION()
        {
            InitializeComponent();
        }

        public void formatInspectNGTable(DataGridView dataGridView1)
        {
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;

            dataGridView1.Columns["INSPECT_TOTAL_QTY"].DefaultCellStyle.Format = "#,0";
            //dataGridView1.Columns["OUT_CFM_QTY"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["INSPECT_OK_QTY"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["INSPECT_TOTAL_LOSS_QTY"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["INSPECT_TOTAL_NG_QTY"].DefaultCellStyle.Format = "#,0";
           




            dataGridView1.Columns["INSPECT_TOTAL_QTY"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["INSPECT_TOTAL_QTY"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["INSPECT_TOTAL_QTY"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["INSPECT_OK_QTY"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["INSPECT_OK_QTY"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["INSPECT_OK_QTY"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["INSPECT_TOTAL_LOSS_QTY"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["INSPECT_TOTAL_LOSS_QTY"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["INSPECT_TOTAL_LOSS_QTY"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["INSPECT_TOTAL_NG_QTY"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["INSPECT_TOTAL_NG_QTY"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["INSPECT_TOTAL_NG_QTY"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);


            dataGridView1.Columns["G_NAME"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.BackColor = Color.Black;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);


            Color loss_color = Color.Purple;
            Color materialNG_color = Color.Brown;
            Color processNG_color = Color.Blue;

            dataGridView1.Columns["ERR1"].HeaderCell.Style.BackColor = loss_color;
            dataGridView1.Columns["ERR2"].HeaderCell.Style.BackColor = loss_color;
            dataGridView1.Columns["ERR3"].HeaderCell.Style.BackColor = loss_color;
            dataGridView1.Columns["ERR4"].HeaderCell.Style.BackColor = materialNG_color;
            dataGridView1.Columns["ERR5"].HeaderCell.Style.BackColor = materialNG_color;
            dataGridView1.Columns["ERR6"].HeaderCell.Style.BackColor = materialNG_color;
            dataGridView1.Columns["ERR7"].HeaderCell.Style.BackColor = materialNG_color;
            dataGridView1.Columns["ERR8"].HeaderCell.Style.BackColor = materialNG_color;
            dataGridView1.Columns["ERR9"].HeaderCell.Style.BackColor = materialNG_color;
            dataGridView1.Columns["ERR10"].HeaderCell.Style.BackColor = materialNG_color;
            dataGridView1.Columns["ERR11"].HeaderCell.Style.BackColor = materialNG_color;
            dataGridView1.Columns["ERR12"].HeaderCell.Style.BackColor = processNG_color;
            dataGridView1.Columns["ERR13"].HeaderCell.Style.BackColor = processNG_color;
            dataGridView1.Columns["ERR14"].HeaderCell.Style.BackColor = processNG_color;
            dataGridView1.Columns["ERR15"].HeaderCell.Style.BackColor = processNG_color;
            dataGridView1.Columns["ERR16"].HeaderCell.Style.BackColor = processNG_color;
            dataGridView1.Columns["ERR17"].HeaderCell.Style.BackColor = processNG_color;
            dataGridView1.Columns["ERR18"].HeaderCell.Style.BackColor = processNG_color;
            dataGridView1.Columns["ERR19"].HeaderCell.Style.BackColor = processNG_color;
            dataGridView1.Columns["ERR20"].HeaderCell.Style.BackColor = processNG_color;
            dataGridView1.Columns["ERR21"].HeaderCell.Style.BackColor = processNG_color;
            dataGridView1.Columns["ERR22"].HeaderCell.Style.BackColor = processNG_color;
            dataGridView1.Columns["ERR23"].HeaderCell.Style.BackColor = processNG_color;
            dataGridView1.Columns["ERR24"].HeaderCell.Style.BackColor = processNG_color;
            dataGridView1.Columns["ERR25"].HeaderCell.Style.BackColor = processNG_color;
            dataGridView1.Columns["ERR26"].HeaderCell.Style.BackColor = processNG_color;
            dataGridView1.Columns["ERR27"].HeaderCell.Style.BackColor = processNG_color;
            dataGridView1.Columns["ERR28"].HeaderCell.Style.BackColor = processNG_color;
            dataGridView1.Columns["ERR29"].HeaderCell.Style.BackColor = processNG_color;
            dataGridView1.Columns["ERR30"].HeaderCell.Style.BackColor = processNG_color;
            dataGridView1.Columns["ERR31"].HeaderCell.Style.BackColor = processNG_color;
            dataGridView1.Columns["ERR32"].HeaderCell.Style.BackColor = Color.Red;

        }


        public void formatInspectBalanceTable_YCSX(DataGridView dataGridView1)
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
            dataGridView1.Columns["DA_KIEM_TRA"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["OK_QTY"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["LOSS_NG_QTY"].DefaultCellStyle.Format = "#,0";
            

            dataGridView1.Columns["LOT_TOTAL_INPUT_QTY_EA"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["LOT_TOTAL_INPUT_QTY_EA"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["LOT_TOTAL_INPUT_QTY_EA"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["LOT_TOTAL_OUTPUT_QTY_EA"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["LOT_TOTAL_OUTPUT_QTY_EA"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["LOT_TOTAL_OUTPUT_QTY_EA"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["INSPECT_BALANCE"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["INSPECT_BALANCE"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["INSPECT_BALANCE"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["G_NAME"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.BackColor = Color.Black;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["PROD_REQUEST_NO"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["PROD_REQUEST_NO"].DefaultCellStyle.BackColor = Color.Black;
            dataGridView1.Columns["PROD_REQUEST_NO"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

        }



        public void formatInspectBalanceTable(DataGridView dataGridView1)
        {
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;

            dataGridView1.Columns["PROD_REQUEST_QTY"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["OUT_CFM_QTY"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["LOT_TOTAL_INPUT_QTY_EA"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["LOT_TOTAL_OUTPUT_QTY_EA"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["INSPECT_BALANCE"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["TOTAL_LOSS"].DefaultCellStyle.Format = "#,0";




            dataGridView1.Columns["LOT_TOTAL_INPUT_QTY_EA"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["LOT_TOTAL_INPUT_QTY_EA"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["LOT_TOTAL_INPUT_QTY_EA"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["LOT_TOTAL_OUTPUT_QTY_EA"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["LOT_TOTAL_OUTPUT_QTY_EA"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["LOT_TOTAL_OUTPUT_QTY_EA"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["INSPECT_BALANCE"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["INSPECT_BALANCE"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["INSPECT_BALANCE"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["TOTAL_LOSS"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["TOTAL_LOSS"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["TOTAL_LOSS"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);


            dataGridView1.Columns["G_NAME"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.BackColor = Color.Black;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["PROD_REQUEST_NO"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["PROD_REQUEST_NO"].DefaultCellStyle.BackColor = Color.Black;
            dataGridView1.Columns["PROD_REQUEST_NO"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

        }




        public void formatInspectOutputTable(DataGridView dataGridView1)
        {
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;

            dataGridView1.Columns["PROD_REQUEST_QTY"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["OUTPUT_QTY_EA"].DefaultCellStyle.Format = "#,0";
          



            dataGridView1.Columns["OUTPUT_QTY_EA"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["OUTPUT_QTY_EA"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["OUTPUT_QTY_EA"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["G_NAME"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.BackColor = Color.Black;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["PROD_REQUEST_NO"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["PROD_REQUEST_NO"].DefaultCellStyle.BackColor = Color.Black;
            dataGridView1.Columns["PROD_REQUEST_NO"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

        }


        public void formatInspectInputTable_YCSX(DataGridView dataGridView1)
        {
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;

            dataGridView1.Columns["PROD_REQUEST_QTY"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["INPUT_QTY_EA"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["INPUT_QTY_KG"].DefaultCellStyle.Format = "#,0.0";



            dataGridView1.Columns["INPUT_QTY_EA"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["INPUT_QTY_EA"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["INPUT_QTY_EA"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["G_NAME"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.BackColor = Color.Black;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["PROD_REQUEST_NO"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["PROD_REQUEST_NO"].DefaultCellStyle.BackColor = Color.Black;
            dataGridView1.Columns["PROD_REQUEST_NO"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

        }



        public void formatInspectInputTable(DataGridView dataGridView1)
        {
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;

            dataGridView1.Columns["PROD_REQUEST_QTY"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["INPUT_QTY_EA"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["INPUT_QTY_KG"].DefaultCellStyle.Format = "#,0.0";
           


            dataGridView1.Columns["INPUT_QTY_EA"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["INPUT_QTY_EA"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["INPUT_QTY_EA"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["G_NAME"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.BackColor = Color.Black;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["PROD_REQUEST_NO"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["PROD_REQUEST_NO"].DefaultCellStyle.BackColor = Color.Black;
            dataGridView1.Columns["PROD_REQUEST_NO"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);



        }


        public void tranhapkiem()
        {
            try
            {
                this.dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();
                dt = pro.report_inspection_all_input_data(generate_condition_inspection_input());



                DataRow drpic = dt.NewRow();
                int colnumdtpic = dt.Columns.Count;
                int rownumdtpic = dt.Rows.Count;
                double sumdtpic = 0.0;
               
                    for (int j = 0; j < rownumdtpic; j++)
                    {
                        sumdtpic += double.Parse(dt.Rows[j][dt.Columns[13]].ToString());
                    }                   
                   
               
                drpic[1] = "TOTAL";                
                drpic[13] = sumdtpic;

                dt.Rows.InsertAt(drpic, 0);


                dataGridView1.DataSource = dt;
                formatInspectInputTable(dataGridView1);


                dataGridView1.Rows[0].DefaultCellStyle.ForeColor = Color.White;
                dataGridView1.Rows[0].DefaultCellStyle.BackColor = Color.Green;
                dataGridView1.Rows[0].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);



                setRowNumber(dataGridView1);
                MessageBox.Show("Đã load : " + dt.Rows.Count + "dòng");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Loi : " + ex.ToString());
            }

        }

        public void traxuatkiem()
        {
            try
            {
                this.dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dt = pro.report_inspection_all_output_data(generate_condition_inspection_output());

                DataRow drpic = dt.NewRow();
                int colnumdtpic = dt.Columns.Count;
                int rownumdtpic = dt.Rows.Count;
                double sumdtpic = 0.0;

                for (int j = 0; j < rownumdtpic; j++)
                {
                    sumdtpic += double.Parse(dt.Rows[j][dt.Columns[13]].ToString());
                }


                drpic[1] = "TOTAL";                
                drpic[13] = sumdtpic;
                dt.Rows.InsertAt(drpic, 0);


                dataGridView1.DataSource = dt;
            formatInspectOutputTable(dataGridView1);

                dataGridView1.Rows[0].DefaultCellStyle.ForeColor = Color.White;
                dataGridView1.Rows[0].DefaultCellStyle.BackColor = Color.Green;
                dataGridView1.Rows[0].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);


                setRowNumber(dataGridView1);
            MessageBox.Show("Đã load : " + dt.Rows.Count + "dòng");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Loi : " + ex.ToString());
            }
        }

        public void trabalancekiem()
        {
            try
            {
                this.dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            // MessageBox.Show(generate_condition_inspection_balance());
            dt = pro.report_inspection_all_balance_data(generate_condition_inspection_balance());
            dataGridView1.DataSource = dt;
            formatInspectBalanceTable(dataGridView1);
            setRowNumber(dataGridView1);
            MessageBox.Show("Đã load : " + dt.Rows.Count + "dòng");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Loi : " + ex.ToString());
            }
}
        public void traNGkiem()
        {
            this.dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            // MessageBox.Show(generate_condition_inspection_balance());
            dt = pro.report_inspection_all_NG_data(generate_condition_inspection_NG());
            dataGridView1.DataSource = dt;
            formatInspectNGTable(dataGridView1);
            setRowNumber(dataGridView1);
            changeHeaderText(dataGridView1);
            MessageBox.Show("Đã load : " + dt.Rows.Count + "dòng");
        }


        public void traINPUTYCSX()
        {
            try
            {
                this.dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();
                dt = pro.report_inspection_all_input_data_YCSX(generate_condition_inspection_input_YCSX());

               

                dataGridView1.DataSource = dt;
                formatInspectInputTable_YCSX(dataGridView1);

              

                setRowNumber(dataGridView1);
                MessageBox.Show("Đã load : " + dt.Rows.Count + "dòng");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Loi : " + ex.ToString());
            }


        }

        public void traOUTPUTYCSX()
        {
            try
            {
                this.dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();
                dt = pro.report_inspection_all_output_data_YCSX(generate_condition_inspection_output_YCSX());
                dataGridView1.DataSource = dt;
                formatInspectOutputTable(dataGridView1);
                setRowNumber(dataGridView1);
                MessageBox.Show("Đã load : " + dt.Rows.Count + "dòng");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Loi : " + ex.ToString());
            }
        }

        public void traINOUTYCSX()
        {
            try
            {
                this.dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();
                // MessageBox.Show(generate_condition_inspection_balance());
                dt = pro.report_inspection_all_balance_data_YCSX(generate_condition_inspection_balance_YCSX());
                DataRow drpic = dt.NewRow();
                int colnumdtpic = dt.Columns.Count;
                int rownumdtpic = dt.Rows.Count;
                double sumin = 0.0, sumout = 0.0, dakiem = 0.0, ok = 0.0, loss = 0.0, kiembalance = 0.0;

                for (int j = 0; j < rownumdtpic; j++)
                {
                    sumin += double.Parse(dt.Rows[j][dt.Columns[8]].ToString());
                    sumout += double.Parse(dt.Rows[j][dt.Columns[9]].ToString());
                    dakiem += double.Parse(dt.Rows[j][dt.Columns[10]].ToString());
                    ok += double.Parse(dt.Rows[j][dt.Columns[11]].ToString());
                    loss += double.Parse(dt.Rows[j][dt.Columns[12]].ToString());
                    kiembalance += double.Parse(dt.Rows[j][dt.Columns[13]].ToString());
                }


                drpic[1] = "TOTAL";

                drpic[8] = sumin;
                drpic[9] = sumout;
                drpic[10] = dakiem;
                drpic[11] = ok;
                drpic[12] = loss;
                drpic[13] = kiembalance;



                dt.Rows.InsertAt(drpic, 0);


                dataGridView1.DataSource = dt;
                formatInspectBalanceTable_YCSX(dataGridView1);
                dataGridView1.Rows[0].DefaultCellStyle.ForeColor = Color.White;
                dataGridView1.Rows[0].DefaultCellStyle.BackColor = Color.Green;
                dataGridView1.Rows[0].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);
                setRowNumber(dataGridView1);
                MessageBox.Show("Đã load : " + dt.Rows.Count + "dòng");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Loi : " + ex.ToString());
            }
        }






        public void changeHeaderText(DataGridView dtg1)
        {
            dtg1.Columns["ERR1"].HeaderText = "Loss thêm túi-포장 로스";
            dtg1.Columns["ERR2"].HeaderText = "Loss bóc đầu cuối-초종 파괴 검사 로스";
            dtg1.Columns["ERR3"].HeaderText = "Loss điểm nối-이음애 로스";
            dtg1.Columns["ERR4"].HeaderText = "Dị vật/chấm gel-원단 이물/겔 점";
            dtg1.Columns["ERR5"].HeaderText = "Nhăn VL-원단 주름";
            dtg1.Columns["ERR6"].HeaderText = "Loang bẩn VL-얼룩";
            dtg1.Columns["ERR7"].HeaderText = "Bóng khí VL-원단 기포";
            dtg1.Columns["ERR8"].HeaderText = "Xước VL-원단 스크래치";
            dtg1.Columns["ERR9"].HeaderText = "Chấm lồi lõm VL-원단 눌림";
            dtg1.Columns["ERR10"].HeaderText = "Keo VL-원단 찐";
            dtg1.Columns["ERR11"].HeaderText = "Lông PE VL-원단 버 (털 모양)";
            dtg1.Columns["ERR12"].HeaderText = "Lỗi IN (Dây mực)-잉크 튐";
            dtg1.Columns["ERR13"].HeaderText = "Lỗi IN (Mất nét)-글자 유실";
            dtg1.Columns["ERR14"].HeaderText = "Lỗi IN (Lỗi màu)-색상 불량";
            dtg1.Columns["ERR15"].HeaderText = "Lỗi IN (Chấm đường khử keo)-점착 제거 선 점 불량";
            dtg1.Columns["ERR16"].HeaderText = "DIECUT (Lệch/Viền màu)-타발 편심";
            dtg1.Columns["ERR17"].HeaderText = "DIECUT (Sâu)-과타발";
            dtg1.Columns["ERR18"].HeaderText = "DIECUT (Nông)-미타발";
            dtg1.Columns["ERR19"].HeaderText = "DIECUT (BAVIA)-타발 버";
            dtg1.Columns["ERR20"].HeaderText = "Mất bước-차수 누락";
            dtg1.Columns["ERR21"].HeaderText = "Xước-스크래치";
            dtg1.Columns["ERR22"].HeaderText = "Nhăn gãy-주름꺽임";
            dtg1.Columns["ERR23"].HeaderText = "Hằn-자국";
            dtg1.Columns["ERR24"].HeaderText = "Sót rác-미 스크랩";
            dtg1.Columns["ERR25"].HeaderText = "Bóng Khí-기포";
            dtg1.Columns["ERR26"].HeaderText = "Bẩn keo bề mặt-표면 찐";
            dtg1.Columns["ERR27"].HeaderText = "Chấm thủng/lồi lõm-찍힘";
            dtg1.Columns["ERR28"].HeaderText = "Bụi trong-내면 이물";
            dtg1.Columns["ERR29"].HeaderText = "Hụt Tape-테이프 줄여듬";
            dtg1.Columns["ERR30"].HeaderText = "Bong keo-찐 벗겨짐";
            dtg1.Columns["ERR31"].HeaderText = "Lấp lỗ sensor-센서 홀 막힘";
            dtg1.Columns["ERR32"].HeaderText = "Marking SX-생산 마킹 구간 썩임";
        }


        private void button1_Click(object sender, EventArgs e)
        {
            tranhapkiem(); 
            label1.Text = "TRA NHẬP KIỂM THEO LOT";

        }

        private void INSPECTION_Load(object sender, EventArgs e)
        {
            dataGridView1.MultiSelect = true;
            //dataGridView1.SelectionMode = DataGridViewSelectionMode.;
            dataGridView1.ReadOnly = true;
            checkBox1.Checked = true;
            if (!System.Windows.Forms.SystemInformation.TerminalServerSession)
            {
                Type dgvType = dataGridView1.GetType();
                PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                  BindingFlags.Instance | BindingFlags.NonPublic);
                pi.SetValue(dataGridView1, true, null);
            }


        }

        private void button2_Click(object sender, EventArgs e)
        {
            traxuatkiem();
            label1.Text = "TRA XUẤT KIỂM THEO LOT";
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //trabalancekiem();
            traINOUTYCSX();
            label1.Text = "TRA NHẬP XUẤT KIỂM THEO YCSX";
        }


        public string generate_condition_inspection_NG()
        {
            string query = "WHERE ";
            string fromdate, todate;
            fromdate = STYMD2(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day);
            todate = STYMD2(dateTimePicker2.Value.Year, dateTimePicker2.Value.Month, dateTimePicker2.Value.Day + 1);
            //MessageBox.Show(fromdate);
            string ngaythang = "ZTBINSPECTNGTB.INSPECT_DATETIME BETWEEN '" + fromdate + " 00:00:00' AND '" + todate + " 23:59:59' ";
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


            string lotno;
            if (textBox5.Text != "")
            {
                lotno = "AND ZTBINSPECTNGTB.PROCESS_LOT_NO= '" + textBox5.Text + "' ";
            }
            else
            {
                lotno = "";
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

            string id = "";
            if (textBox6.Text != "")
            {
                id = "AND INSPECT_ID=" + textBox6.Text;
            }
            else
            {
                id = "";
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

            query += ngaythang + code + cust_name_kd + id + cmscode + ycsxno + lotno;
            return query;
        }




        public string generate_condition_inspection_output_YCSX()
        {
            string query = "WHERE 1=1";
            

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

            query +=  code + cust_name_kd +  cmscode + ycsxno ;
            return query;
        }



        public string generate_condition_inspection_input_YCSX()
        {
            string query = "WHERE 1=1";
           
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

            query +=  code + cust_name_kd +  cmscode + ycsxno ;
            return query;
        }




        public string generate_condition_inspection_input()
        {
            string query = "WHERE ";
            string fromdate, todate;
            fromdate = STYMD2(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day);
            todate = STYMD2(dateTimePicker2.Value.Year, dateTimePicker2.Value.Month, dateTimePicker2.Value.Day);
            //MessageBox.Show(fromdate);
            string ngaythang = "ZTBINSPECTINPUTTB.INPUT_DATETIME BETWEEN '" + fromdate + " 00:00:00' AND '" + todate + " 23:59:59' ";
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


            string lotno;
            if (textBox5.Text != "")
            {
                lotno = "AND P501_A.PROCESS_LOT_NO= '" + textBox5.Text + "' ";
            }
            else
            {
                lotno = "";
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
           
            string id = "";
            if (textBox6.Text != "")
            {
                id = "AND INSPECT_INPUT_ID=" + textBox6.Text;
            }
            else
            {
                id = "";
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

            query += ngaythang + code +  cust_name_kd +  id + cmscode + ycsxno + lotno;
            return query;
        }



        public string generate_condition_inspection_output()
        {
            string query = "WHERE ";
            string fromdate, todate;
            fromdate = STYMD2(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day);
            todate = STYMD2(dateTimePicker2.Value.Year, dateTimePicker2.Value.Month, dateTimePicker2.Value.Day);
            //MessageBox.Show(fromdate);
            string ngaythang = "ZTBINSPECTOUTPUTTB.OUTPUT_DATETIME BETWEEN '" + fromdate + " 00:00:00' AND '" + todate + " 23:59:59' ";
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


            string lotno;
            if (textBox5.Text != "")
            {
                lotno = "AND P501_A.PROCESS_LOT_NO= '" + textBox5.Text + "' ";
            }
            else
            {
                lotno = "";
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

            string id = "";
            if (textBox6.Text != "")
            {
                id = "AND INSPECT_OUTPUT_ID=" + textBox6.Text;
            }
            else
            {
                id = "";
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

            query += ngaythang + code + cust_name_kd + id + cmscode + ycsxno + lotno;
            return query;
        }

        public string generate_condition_inspection_balance_YCSX()
        {
            string query = "WHERE 1=1";

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

            query += code + cust_name_kd + cmscode + ycsxno;
            return query;
        }

        public string generate_condition_inspection_balance()
        {
            string query = "WHERE 1=1";
           
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


            string lotno;
            if (textBox5.Text != "")
            {
                lotno = "AND P501.PROCESS_LOT_NO= '" + textBox5.Text + "' ";
            }
            else
            {
                lotno = "";
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

            query +=  code +  cust_name_kd +  cmscode + ycsxno + lotno;
            return query;
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

        private void button4_Click(object sender, EventArgs e)
        {
            //MessageBox.Show(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            Form inspection_input = new INSPECT_INPUT();
            inspection_input.Show();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Form inspection_output = new INSPECT_OUTPUT();
            inspection_output.Show();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            this.dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            // MessageBox.Show(generate_condition_inspection_balance());
            dt = pro.report_inspection_all_balance_data(generate_condition_inspection_balance());
            dataGridView1.DataSource = dt;
            formatInspectBalanceTable(dataGridView1);
            setRowNumber(dataGridView1);
            MessageBox.Show("Đã load : " + dt.Rows.Count + "dòng");
        }

        private void nHẬPKIỂMToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //MessageBox.Show(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            Form inspection_input = new INSPECT_INPUT();
            inspection_input.Show();
        }

        private void xUẤTKIỂMToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form inspection_output = new INSPECT_OUTPUT();
            inspection_output.Show();
        }

        private void nGKIỂMToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form inspection_ng = new INSPECT_NG();
            inspection_ng.Show();
           
        }

        private void tRANHẬPKIỂMToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tranhapkiem();
            label1.Text = "TRA NHẬP KIỂM THEO LOT";
        }

        private void tRAXUẤTKIỂMToolStripMenuItem_Click(object sender, EventArgs e)
        {
            traxuatkiem();
            label1.Text = "TRA XUẤT KIỂM THEO LOT";
        }

        private void tRANGDATAToolStripMenuItem_Click(object sender, EventArgs e)
        {
            traNGkiem();
            label1.Text = "TRA NG KIỂM THEO NHẬT KÝ";
        }

        private void tRATỒNKIỂMTHEOLOTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            trabalancekiem();
            label1.Text = "TRA TỒN KIỂM THEO LOT";
        }

        private void tRATỒNKIỂMTHEOYCSXToolStripMenuItem_Click(object sender, EventArgs e)
        {
          
        }

        private void tRAINPUTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            traINPUTYCSX();
            label1.Text = "TRA NHẬP KIỂM THEO YCSX";
        }

        private void tRAOUTPUTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            traOUTPUTYCSX();
            label1.Text = "TRA XUẤT KIỂM THEO YCSX";
        }

        private void tRAINOUTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            traINOUTYCSX();
            label1.Text = "TRA NHẬP XUẤT KIỂM THEO YCSX";
        }
    }
}
