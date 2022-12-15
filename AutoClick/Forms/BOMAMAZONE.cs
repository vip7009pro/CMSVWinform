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
    public partial class BOMAMAZONE : Form
    {
        public string Login_ID = "";
        public BOMAMAZONE()
        {
            InitializeComponent();
        }

        public void formatDataGridViewAllCode(DataGridView dataGridView1)
        {
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;          
      
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.BackColor = Color.Brown;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.ForeColor = Color.Yellow;

            dataGridView1.Columns["G_CODE"].DefaultCellStyle.BackColor = Color.Brown;
            dataGridView1.Columns["G_CODE"].DefaultCellStyle.ForeColor = Color.Yellow;
        }

        public void formatDataGridViewBOMAmazone(DataGridView dataGridView1)
        {
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;

            dataGridView1.Columns["G_NAME"].DefaultCellStyle.BackColor = Color.Yellow;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.ForeColor = Color.Black;

            dataGridView1.Columns["G_CODE"].DefaultCellStyle.BackColor = Color.Brown;
            dataGridView1.Columns["G_CODE"].DefaultCellStyle.ForeColor = Color.Yellow;

            dataGridView1.Columns["GIATRI"].DefaultCellStyle.BackColor = Color.Yellow;
            dataGridView1.Columns["GIATRI"].DefaultCellStyle.ForeColor = Color.Black;

            dataGridView1.Columns["REMARK"].DefaultCellStyle.BackColor = Color.Brown;
            dataGridView1.Columns["REMARK"].DefaultCellStyle.ForeColor = Color.Yellow;
        }


        private void button1_Click(object sender, EventArgs e)
        {
            ProductBLL pro = new ProductBLL();
            DataTable dt = pro.getMcodeInfo(textBox1.Text);
            dataGridView2.DataSource = dt;
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                ProductBLL pro = new ProductBLL();
                DataTable dt = pro.getAmazoneBOM(textBox1.Text);
                dataGridView2.DataSource = dt;
            }
        }

        private void dataGridView2_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.Columns.Count > 0)
            {
                dataGridView1.DataSource = null;
                if(dataGridView1.Columns.Contains("G_CODE") == true) dataGridView1.Columns.Remove("G_CODE");
                if (dataGridView1.Columns.Contains("G_NAME") == true) dataGridView1.Columns.Remove("G_NAME");
                if (dataGridView1.Columns.Contains("GIATRI") == true) dataGridView1.Columns.Remove("GIATRI");
                if (dataGridView1.Columns.Contains("REMARK") == true) dataGridView1.Columns.Remove("REMARK");
                              
            }
            DataGridViewRow row = dataGridView2.Rows[e.RowIndex];
            string G_CODE = row.Cells["G_CODE"].Value.ToString();
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dt = pro.checkBOMAmazone(G_CODE);
            dataGridView1.DataSource = null;
            dataGridView1.DataSource = dt;
            formatDataGridViewBOMAmazone(dataGridView1);


        }

        private void BOMAMAZONE_Load(object sender, EventArgs e)
        {
            if (!System.Windows.Forms.SystemInformation.TerminalServerSession)
            {
                Type dgvType = dataGridView1.GetType();
                PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                  BindingFlags.Instance | BindingFlags.NonPublic);
                pi.SetValue(dataGridView1, true, null);
            }

            if (!System.Windows.Forms.SystemInformation.TerminalServerSession)
            {
                Type dgvType = dataGridView2.GetType();
                PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                  BindingFlags.Instance | BindingFlags.NonPublic);
                pi.SetValue(dataGridView2, true, null);
            }

            if (!System.Windows.Forms.SystemInformation.TerminalServerSession)
            {
                Type dgvType = dataGridView3.GetType();
                PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                  BindingFlags.Instance | BindingFlags.NonPublic);
                pi.SetValue(dataGridView3, true, null);
            }
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;

            dataGridView3.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();

            comboBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            comboBox1.AutoCompleteSource = AutoCompleteSource.ListItems;

            dt = pro.getamazonedesign();
            comboBox1.DataSource = dt;
            comboBox1.DisplayMember = "G_NAME";
            comboBox1.ValueMember = "G_CODE";
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ProductBLL pro = new ProductBLL();
            try
            {
                bool loi = false;
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (!row.IsNewRow)
                    {
                        string GIATRI = row.Cells["GIATRI"].Value.ToString();
                        if (GIATRI.Length > 0)
                        {
                            if (GIATRI[0].ToString() == " " || GIATRI[GIATRI.Length - 1].ToString() == " ")
                            {
                                dataGridView1.Rows[row.Index].Cells["GIATRI"].Style.BackColor = Color.Red;
                                loi = true;
                            }
                        }
                           
                    }
                }

                if(loi == true)
                {
                    MessageBox.Show("Giá trị không được chứa dấu cách ở 2 đầu");
                }
                else
                {
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (!row.IsNewRow)
                        {
                            string
                                G_CODE = row.Cells["G_CODE"].Value.ToString(),
                                G_CODE_MAU = row.Cells["G_CODE_MAU"].Value.ToString(),
                                DOITUONG_NO = row.Cells["DOITUONG_NO"].Value.ToString(),
                                GIATRI = row.Cells["GIATRI"].Value.ToString(),
                                REMARK = row.Cells["REMARK"].Value.ToString(),
                                EMPL_NO = Login_ID;
                            string insertValueBOMAmazone = $"('002','{G_CODE}','{G_CODE_MAU}','{DOITUONG_NO}', '{GIATRI}','{REMARK}',GETDATE(), '{EMPL_NO}', GETDATE(), '{EMPL_NO}')"; 
                            pro.InsertBOMAmazone(insertValueBOMAmazone);
                        }
                    }
                    MessageBox.Show("Thêm BOM AMAZONE thành công !");

                }

                

            }
            catch (Exception ex)
            {

                MessageBox.Show("Lỗi: + " + ex.ToString());
            }
            

        }

        private void button2_Click(object sender, EventArgs e)
        {
            ProductBLL pro = new ProductBLL();
            try
            {
                bool loi = false;
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (!row.IsNewRow)
                    {
                        string GIATRI = row.Cells["GIATRI"].Value.ToString();
                        if(GIATRI.Length>0)
                        {
                            if (GIATRI[0].ToString() == " " || GIATRI[GIATRI.Length - 1].ToString() == " ")
                            {
                                dataGridView1.Rows[row.Index].Cells["GIATRI"].Style.BackColor = Color.Red;
                                loi = true;
                            }

                        }
                       
                    }
                }

                if (loi == true)
                {
                    MessageBox.Show("Giá trị không được chứa dấu cách ở 2 đầu");
                }
                else
                {
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (!row.IsNewRow)
                        {
                            string
                                G_CODE = row.Cells["G_CODE"].Value.ToString(),
                                G_CODE_MAU = row.Cells["G_CODE_MAU"].Value.ToString(),
                                DOITUONG_NO = row.Cells["DOITUONG_NO"].Value.ToString(),
                                GIATRI = row.Cells["GIATRI"].Value.ToString(),
                                REMARK = row.Cells["REMARK"].Value.ToString(),
                                EMPL_NO = Login_ID;
                            string updateValue = $" GIATRI = '{GIATRI}', REMARK='{REMARK}', UPD_DATE=GETDATE(), UPD_EMPL='{EMPL_NO}' WHERE G_CODE='{G_CODE}' AND G_CODE_MAU='{G_CODE_MAU}' AND DOITUONG_NO='{DOITUONG_NO}'";
                            pro.UpdateBOMAmazone(updateValue);
                        }
                    }
                    MessageBox.Show("Update BOM AMAZONE thành công !");
                }                
            }
            catch (Exception ex)
            {

                MessageBox.Show("Lỗi: + " + ex.ToString());
            }

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                ProductBLL pro = new ProductBLL();
                DataTable dt = pro.getMcodeInfo(textBox2.Text);
                dataGridView3.DataSource = dt;
                formatDataGridViewAllCode(dataGridView3);
            }
        }

        private void dataGridView3_RowEnter(object sender, DataGridViewCellEventArgs e)
        {

            try
            {
                if (dataGridView1.Columns.Count > 0)
                {
                    dataGridView1.DataSource = null;
                    if (dataGridView1.Columns.Contains("G_CODE") == true) dataGridView1.Columns.Remove("G_CODE");
                    if (dataGridView1.Columns.Contains("G_NAME") == true) dataGridView1.Columns.Remove("G_NAME");
                    if (dataGridView1.Columns.Contains("GIATRI") == true) dataGridView1.Columns.Remove("GIATRI");
                    if (dataGridView1.Columns.Contains("REMARK") == true) dataGridView1.Columns.Remove("REMARK");
                }



                string G_CODE_MAU = comboBox1.SelectedValue.ToString();
                DataGridViewRow row = dataGridView3.SelectedRows[0];
                string G_CODE = row.Cells["G_CODE"].Value.ToString();
                string G_NAME = row.Cells["G_NAME"].Value.ToString();
                //MessageBox.Show(G_NAME + " | " + G_CODE);
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();
                dt = pro.getamazonedesign2(G_CODE_MAU);
                dataGridView1.DataSource = dt;

                DataGridViewColumn G_NAME_COL = new DataGridViewColumn();
                G_NAME_COL.Name = "G_NAME";
                G_NAME_COL.HeaderText = "G_NAME";
                G_NAME_COL.Width = 90;
                G_NAME_COL.ReadOnly = false;
                G_NAME_COL.CellTemplate = new DataGridViewTextBoxCell();
                dataGridView1.Columns.Insert(0, G_NAME_COL);
                for (int r = 0; r < dataGridView1.Rows.Count; r++)
                {
                    dataGridView1.Rows[r].Cells["G_NAME"].Value = G_NAME;
                }

                DataGridViewColumn G_CODE_COL = new DataGridViewColumn();
                G_CODE_COL.Name = "G_CODE";
                G_CODE_COL.HeaderText = "G_CODE";
                G_CODE_COL.Width = 90;
                G_CODE_COL.ReadOnly = false;
                G_CODE_COL.CellTemplate = new DataGridViewTextBoxCell();
                dataGridView1.Columns.Insert(0, G_CODE_COL);
                for (int r = 0; r < dataGridView1.Rows.Count; r++)
                {
                    dataGridView1.Rows[r].Cells["G_CODE"].Value = G_CODE;
                }


                DataGridViewColumn GIATRI_COL = new DataGridViewColumn();
                GIATRI_COL.Name = "GIATRI";
                GIATRI_COL.HeaderText = "GIATRI";
                GIATRI_COL.Width = 90;
                GIATRI_COL.ReadOnly = false;
                GIATRI_COL.CellTemplate = new DataGridViewTextBoxCell();
                dataGridView1.Columns.Insert(6, GIATRI_COL);
                for (int r = 0; r < dataGridView1.Rows.Count; r++)
                {
                    dataGridView1.Rows[r].Cells["GIATRI"].Value = "";
                }

                DataGridViewColumn REMARK_COL = new DataGridViewColumn();
                REMARK_COL.Name = "REMARK";
                REMARK_COL.HeaderText = "REMARK";
                REMARK_COL.Width = 90;
                REMARK_COL.ReadOnly = false;
                REMARK_COL.CellTemplate = new DataGridViewTextBoxCell();
                dataGridView1.Columns.Insert(7, REMARK_COL);
                for (int r = 0; r < dataGridView1.Rows.Count; r++)
                {
                    dataGridView1.Rows[r].Cells["REMARK"].Value = "";
                }

                formatDataGridViewBOMAmazone(dataGridView1);




            }
            catch (Exception ex)
            {

            }

        }
    }
}
