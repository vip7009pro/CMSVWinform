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
    public partial class NewCodeBom : Form
    {
        public string EMPL_NO = "";
        public NewCodeBom()
        {
            InitializeComponent();
        }

        private void gradientPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        public void changeInputStatus(bool status)
        {
            textBox1.Enabled = status;
            textBox2.Enabled = status;
            textBox3.Enabled = status;
            textBox4.Enabled = !status;
            textBox5.Enabled = status;
            textBox6.Enabled = status;
            textBox7.Enabled = status;
            textBox8.Enabled = status;
            textBox9.Enabled = status;
            textBox10.Enabled = status;
            textBox11.Enabled = status;
            textBox12.Enabled = status;
            textBox13.Enabled = status;
            textBox14.Enabled = status;
            textBox15.Enabled = status;
            textBox16.Enabled = status;
            textBox17.Enabled = status;
            textBox18.Enabled = status;
            textBox19.Enabled = status;
            textBox21.Enabled = status;
            textBox22.Enabled = status;
            textBox23.Enabled = status;
            textBox24.Enabled = status;
            textBox29.Enabled = status;
            textBox30.Enabled = status;

            richTextBox1.Enabled = status;
            comboBox1.Enabled = status;
            comboBox2.Enabled = status;
            comboBox3.Enabled = status;
            comboBox4.Enabled = status;
            comboBox5.Enabled = status;
            comboBox6.Enabled = status;
            comboBox7.Enabled = status;
            comboBox8.Enabled = status;
            comboBox9.Enabled = status;
            checkBox1.Enabled = status;
            checkBox2.Enabled = status;
            dataGridView2.Enabled = !status;
            dataGridView1.ReadOnly = !status;
            radioButton1.Enabled = status;
            radioButton2.Enabled = status;
            radioButton3.Enabled = status;
            button3.Enabled = status;
            button2.Enabled = status;
            button4.Enabled = status;
            button5.Enabled = status;
            button6.Enabled = status;
            button7.Enabled = status;            
            button10.Enabled = status;
            button11.Enabled = status;
            if (status)
            {                
                label43.Text = "Bật sửa";
            }
            else
            {               
                label43.Text = "Khóa sửa";
            }
        }

        public void loadFormFunction()
        {
            if (!System.Windows.Forms.SystemInformation.TerminalServerSession)
            {
                Type dgvType = dataGridView2.GetType();
                PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                  BindingFlags.Instance | BindingFlags.NonPublic);
                pi.SetValue(dataGridView2, true, null);
            }

            comboBox7.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            comboBox7.AutoCompleteSource = AutoCompleteSource.ListItems;

            comboBox9.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            comboBox9.AutoCompleteSource = AutoCompleteSource.ListItems;

            ProductBLL pro = new ProductBLL();
            DataTable dt = pro.report_getCustomerList();

            comboBox7.DataSource = dt;
            comboBox7.ValueMember = "CUST_CD";
            comboBox7.DisplayMember = "CUST_NAME_KD";

            dt = pro.report_MaterialList();
            comboBox9.DataSource = dt;
            comboBox9.ValueMember = "M_CODE";
            comboBox9.DisplayMember = "M_NAME_SIZEZ";
            radioButton1.Checked = true;

            changeInputStatus(false);


        }
        private void NewCodeBom_Load(object sender, EventArgs e)
        {

            // var task1 = Task.Factory.StartNew(loadFormFunction);
            loadFormFunction();

        }

        public void updateCODEANDBOM()
        {
            CMSCODE newcode = new CMSCODE();
            newcode.cust_cd = comboBox7.SelectedValue.ToString();
            newcode.project = textBox2.Text;
            newcode.model = textBox3.Text;
            newcode.dactinh = comboBox8.SelectedIndex.ToString();
            newcode.phanloai = comboBox1.SelectedIndex.ToString();
            newcode.prod_type = comboBox1.SelectedItem.ToString();
            newcode.g_name_kd = textBox5.Text;
            newcode.description = textBox6.Text;
            newcode.main_material = textBox7.Text;
            newcode.g_name = textBox8.Text;
            newcode.length = textBox16.Text;
            newcode.width = textBox15.Text;
            newcode.feeding = textBox14.Text;
            newcode.cavity_hang = textBox13.Text;
            newcode.cavity_cot = textBox12.Text;
            newcode.k_c_hang = textBox11.Text;
            newcode.k_c_cot = textBox10.Text;
            newcode.k_c_liner_trai = textBox9.Text;
            newcode.roll_open_direction = comboBox2.SelectedIndex.ToString();
            newcode.k_c_liner_phai = textBox17.Text;
            newcode.knife = comboBox4.SelectedIndex.ToString();
            newcode.knife_lifecycle = textBox19.Text;
            newcode.knife_price = textBox18.Text;
            newcode.packing_type = comboBox3.SelectedIndex.ToString();
            newcode.packing_qty = textBox21.Text;
            newcode.rpm = textBox22.Text;
            newcode.pin_distance = textBox23.Text;
            newcode.process_type = textBox24.Text;
            newcode.eq1 = comboBox5.Text;
            newcode.eq2 = comboBox6.Text;
            newcode.steps = textBox30.Text;
            newcode.print_times = textBox29.Text;
            newcode.print_yn = checkBox2.Checked.ToString();
            newcode.draw_path = textBox1.Text;
            newcode.use_yn = checkBox1.Checked.ToString();
            newcode.remark = richTextBox1.Text;
            newcode.ins_empl = EMPL_NO;
            newcode.bom = dataGridView1;
            newcode.old_g_code = label38.Text;  
            
            try
            {
                if(newcode.updateCode()  && newcode.checkInput()=="")
                {
                    MessageBox.Show("Update thông tin code hoàn thành");
                    //richTextBox1.Text = newcode.remark;
                }
                else
                {
                    MessageBox.Show("Lỗi: " + newcode.checkInput());
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi ngoại lệ: " + ex.ToString());
            }
             
            //richTextBox1.Text = newcode.remark;
            //MessageBox.Show(newcode.print_yn);
        }


        public void createNewCode()
        {
            CMSCODE newcode = new CMSCODE();
            newcode.cust_cd = comboBox7.SelectedValue.ToString();
            newcode.project = textBox2.Text;
            newcode.model = textBox3.Text;
            newcode.dactinh = comboBox8.SelectedIndex.ToString();
            newcode.phanloai = comboBox1.SelectedIndex.ToString();
            newcode.prod_type = comboBox1.SelectedItem.ToString();
            newcode.g_name_kd = textBox5.Text;
            newcode.description = textBox6.Text;
            newcode.main_material = textBox7.Text;
            newcode.g_name = textBox8.Text;
            newcode.length = textBox16.Text;
            newcode.width = textBox15.Text;
            newcode.feeding = textBox14.Text;
            newcode.cavity_hang = textBox13.Text;
            newcode.cavity_cot = textBox12.Text;
            newcode.k_c_hang = textBox11.Text;
            newcode.k_c_cot = textBox10.Text;
            newcode.k_c_liner_trai = textBox9.Text;
            newcode.roll_open_direction = comboBox2.SelectedIndex.ToString();
            newcode.k_c_liner_phai = textBox17.Text;
            newcode.knife = comboBox4.SelectedIndex.ToString();
            newcode.knife_lifecycle = textBox19.Text;
            newcode.knife_price = textBox18.Text;
            newcode.packing_type = comboBox3.SelectedIndex.ToString();
            newcode.packing_qty = textBox21.Text;
            newcode.rpm = textBox22.Text;
            newcode.pin_distance = textBox23.Text;
            newcode.process_type = textBox24.Text;
            newcode.eq1 = comboBox5.Text;
            newcode.eq2 = comboBox6.Text;
            newcode.steps = textBox30.Text;
            newcode.print_times = textBox29.Text;
            newcode.print_yn = checkBox2.Checked.ToString();
            newcode.draw_path = textBox1.Text;
            newcode.use_yn = checkBox1.Checked.ToString();
            newcode.remark = richTextBox1.Text;
            newcode.ins_empl = EMPL_NO;
            newcode.bom = dataGridView1;
            
            try
            {
                if (newcode.checkInput() == "" && newcode.checkBOM()=="")
                {
                    if (newcode.insertOldCMS())
                    {
                        try
                        {
                            newcode.insertOldBOM();
                            newcode.insertNewBOM();
                            pro.insertMaterialfromBOMtoMTable();
                            MessageBox.Show("Insert code hoàn thành, code mới là :" + newcode.g_code);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Lỗi M140, BOM2: " + ex.ToString());
                        }
                    }
                }
                else
                {
                    MessageBox.Show(newcode.checkInput() + " | " + newcode.checkBOM());

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            //richTextBox1.Text = newcode.remark;
            //MessageBox.Show(newcode.print_yn);
        }

        public void addNewVer()
        {
            CMSCODE newcode = new CMSCODE();
            newcode.cust_cd = comboBox7.SelectedValue.ToString();
            newcode.project = textBox2.Text;
            newcode.model = textBox3.Text;
            newcode.dactinh = comboBox8.SelectedIndex.ToString();
            newcode.phanloai = comboBox1.SelectedIndex.ToString();
            newcode.prod_type = comboBox1.SelectedItem.ToString();
            newcode.g_name_kd = textBox5.Text;
            newcode.description = textBox6.Text;
            newcode.main_material = textBox7.Text;
            newcode.g_name = textBox8.Text;
            newcode.length = textBox16.Text;
            newcode.width = textBox15.Text;
            newcode.feeding = textBox14.Text;
            newcode.cavity_hang = textBox13.Text;
            newcode.cavity_cot = textBox12.Text;
            newcode.k_c_hang = textBox11.Text;
            newcode.k_c_cot = textBox10.Text;
            newcode.k_c_liner_trai = textBox9.Text;
            newcode.roll_open_direction = comboBox2.SelectedIndex.ToString();
            newcode.k_c_liner_phai = textBox17.Text;
            newcode.knife = comboBox4.SelectedIndex.ToString();
            newcode.knife_lifecycle = textBox19.Text;
            newcode.knife_price = textBox18.Text;
            newcode.packing_type = comboBox3.SelectedIndex.ToString();
            newcode.packing_qty = textBox21.Text;
            newcode.rpm = textBox22.Text;
            newcode.pin_distance = textBox23.Text;
            newcode.process_type = textBox24.Text;
            newcode.eq1 = comboBox5.Text;
            newcode.eq2 = comboBox6.Text;
            newcode.steps = textBox30.Text;
            newcode.print_times = textBox29.Text;
            newcode.print_yn = checkBox2.Checked.ToString();
            newcode.draw_path = textBox1.Text;
            newcode.use_yn = checkBox1.Checked.ToString();
            newcode.remark = richTextBox1.Text;
            newcode.ins_empl = EMPL_NO;
            newcode.old_g_code = label38.Text;
            newcode.bom = dataGridView1;

            try
            {

                if (newcode.checkInput() == "")
                {
                    if (newcode.insertVerCMS() && newcode.checkBOM()=="")
                    {
                        try
                        {
                            newcode.insertOldBOM();
                            newcode.insertNewBOM();
                            MessageBox.Show("Thêm version hoàn thành, code ver mới là :" + newcode.g_code);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Lỗi M140, BOM2: " + ex.ToString());
                        }
                    }
                    else
                    {
                        MessageBox.Show("newcode checkbom " + newcode.checkBOM());
                    }
                }
                else
                {
                    MessageBox.Show(newcode.checkInput());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            //richTextBox1.Text = newcode.remark;
            //MessageBox.Show(newcode.print_yn);
        }

        public void addBOM2ofOldCode()
        {
            CMSCODE newcode = new CMSCODE();
            newcode.cust_cd = comboBox7.SelectedValue.ToString();
            newcode.project = textBox2.Text;
            newcode.model = textBox3.Text;
            newcode.dactinh = comboBox8.SelectedIndex.ToString();
            newcode.phanloai = comboBox1.SelectedIndex.ToString();
            newcode.g_name_kd = textBox5.Text;
            newcode.description = textBox6.Text;
            newcode.main_material = textBox7.Text;
            newcode.g_name = textBox8.Text;
            newcode.length = textBox16.Text;
            newcode.width = textBox15.Text;
            newcode.feeding = textBox14.Text;
            newcode.cavity_hang = textBox13.Text;
            newcode.cavity_cot = textBox12.Text;
            newcode.k_c_hang = textBox11.Text;
            newcode.k_c_cot = textBox10.Text;
            newcode.k_c_liner_trai = textBox9.Text;
            newcode.roll_open_direction = comboBox2.SelectedIndex.ToString();
            newcode.k_c_liner_phai = textBox17.Text;
            newcode.knife = comboBox4.SelectedIndex.ToString();
            newcode.knife_lifecycle = textBox19.Text;
            newcode.knife_price = textBox18.Text;
            newcode.packing_type = comboBox3.SelectedIndex.ToString();
            newcode.packing_qty = textBox21.Text;
            newcode.rpm = textBox22.Text;
            newcode.pin_distance = textBox23.Text;
            newcode.process_type = textBox24.Text;
            newcode.eq1 = comboBox5.Text;
            newcode.eq2 = comboBox6.Text;
            newcode.steps = textBox30.Text;
            newcode.print_times = textBox29.Text;
            newcode.print_yn = checkBox2.Checked.ToString();
            newcode.draw_path = textBox1.Text;
            newcode.use_yn = checkBox1.Checked.ToString();
            newcode.remark = richTextBox1.Text;
            newcode.ins_empl = EMPL_NO;
            newcode.old_g_code = label38.Text;
            newcode.bom = dataGridView1;
            try
            {
                ProductBLL pro = new ProductBLL();
                
                if ((pro.checkBOM2(newcode.old_g_code,newcode.old_g_code.Substring(7,1)) == 0) && !(newcode.old_g_code.Length <8) && newcode.checkBOM()=="")
                {
                    if(newcode.checkInput()=="")
                    {
                        try
                        {
                            newcode.insertBOM2OldCode();
                            if (pro.checkBOMSXExist(newcode.old_g_code, newcode.old_g_code.Substring(7, 1)) == 0) 
                            { 
                                newcode.insertOldBOM();
                            }
                            //neu chua co bom sx thi them bom sx luon
                            pro.insertMaterialfromBOMtoMTable();
                            MessageBox.Show("Thêm BOM tính giá hoàn thành cho code:" + newcode.old_g_code);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Lỗi M140, BOM2: " + ex.ToString());
                        }

                    }
                    else
                    {
                        MessageBox.Show("Update đầy đủ thông tin code đang bị trống trước rồi hãy thêm BOM giá, để có thể tính giá");
                    }
                   
                }
                else
                {
                    MessageBox.Show("Đã có bom tính giá của code này rồi, ko thêm được nữa");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public void justCheckBOM()
        {
            CMSCODE newcode = new CMSCODE();
            newcode.cust_cd = comboBox7.SelectedValue.ToString();
            newcode.project = textBox2.Text;
            newcode.model = textBox3.Text;
            newcode.dactinh = comboBox8.SelectedIndex.ToString();
            newcode.phanloai = comboBox1.SelectedIndex.ToString();
            newcode.g_name_kd = textBox5.Text;
            newcode.description = textBox6.Text;
            newcode.main_material = textBox7.Text;
            newcode.g_name = textBox8.Text;
            newcode.length = textBox16.Text;
            newcode.width = textBox15.Text;
            newcode.feeding = textBox14.Text;
            newcode.cavity_hang = textBox13.Text;
            newcode.cavity_cot = textBox12.Text;
            newcode.k_c_hang = textBox11.Text;
            newcode.k_c_cot = textBox10.Text;
            newcode.k_c_liner_trai = textBox9.Text;
            newcode.roll_open_direction = comboBox2.SelectedIndex.ToString();
            newcode.k_c_liner_phai = textBox17.Text;
            newcode.knife = comboBox4.SelectedIndex.ToString();
            newcode.knife_lifecycle = textBox19.Text;
            newcode.knife_price = textBox18.Text;
            newcode.packing_type = comboBox3.SelectedIndex.ToString();
            newcode.packing_qty = textBox21.Text;
            newcode.rpm = textBox22.Text;
            newcode.pin_distance = textBox23.Text;
            newcode.process_type = textBox24.Text;
            newcode.eq1 = comboBox5.Text;
            newcode.eq2 = comboBox6.Text;
            newcode.steps = textBox30.Text;
            newcode.print_times = textBox29.Text;
            newcode.print_yn = checkBox2.Checked.ToString();
            newcode.draw_path = textBox1.Text;
            newcode.use_yn = checkBox1.Checked.ToString();
            newcode.remark = richTextBox1.Text;
            newcode.ins_empl = EMPL_NO;
            newcode.old_g_code = label38.Text;
            newcode.bom = dataGridView1;
            try
            {
                newcode.checkBOM();                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public void cloneBOM()
        {     
            try
            {
                if (dataGridView3.Rows.Count > 0)
                {
                    int seq = 1;
                    ProductBLL pro = new ProductBLL();
                    DataTable dt = new DataTable();
                    foreach (DataGridViewRow row in dataGridView3.Rows)
                    {
                        if (!row.IsNewRow)
                        {

                            string M_NAME = row.Cells["M_NAME"].Value.ToString();
                            string M_CODE = row.Cells["M_CODE"].Value.ToString();
                            string WIDTH_CD = row.Cells["WIDTH_CD"].Value.ToString();
                            string M_QTY = row.Cells["M_QTY"].Value.ToString();

                            dt = pro.getMATInfo(M_NAME);
                            int countrow = dt.Rows.Count;                           

                            string ssprice = (countrow >0 ? dt.Rows[0]["SSPRICE"].ToString(): "");
                            string cmsprice = (countrow > 0 ? dt.Rows[0]["CMSPRICE"].ToString() : "");
                            string slittingprice = (countrow > 0 ? dt.Rows[0]["SLITTING_PRICE"].ToString() : "");
                            string masterwidth = (countrow > 0 ? dt.Rows[0]["MASTER_WIDTH"].ToString() : "");
                            string roll_length = (countrow > 0 ? dt.Rows[0]["ROLL_LENGTH"].ToString() : "");
                            string vendor = (countrow > 0 ? dt.Rows[0]["CUST_CD"].ToString() : "");

                            dataGridView1.Rows.Add(M_CODE, "LIEU",M_NAME, "", WIDTH_CD, "", vendor, "", cmsprice, ssprice, slittingprice, masterwidth, roll_length, M_QTY, "", "");

                        }
                    }
                }
                else
                {
                    MessageBox.Show("Không có dòng nào được clone");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }




        public void updateBOM2ofOldCode()
        {
            CMSCODE newcode = new CMSCODE();
            newcode.cust_cd = comboBox7.SelectedValue.ToString();
            newcode.project = textBox2.Text;
            newcode.model = textBox3.Text;
            newcode.dactinh = comboBox8.SelectedIndex.ToString();
            newcode.phanloai = comboBox1.SelectedIndex.ToString();
            newcode.g_name_kd = textBox5.Text;
            newcode.description = textBox6.Text;
            newcode.main_material = textBox7.Text;
            newcode.g_name = textBox8.Text;
            newcode.length = textBox16.Text;
            newcode.width = textBox15.Text;
            newcode.feeding = textBox14.Text;
            newcode.cavity_hang = textBox13.Text;
            newcode.cavity_cot = textBox12.Text;
            newcode.k_c_hang = textBox11.Text;
            newcode.k_c_cot = textBox10.Text;
            newcode.k_c_liner_trai = textBox9.Text;
            newcode.roll_open_direction = comboBox2.SelectedIndex.ToString();
            newcode.k_c_liner_phai = textBox17.Text;
            newcode.knife = comboBox4.SelectedIndex.ToString();
            newcode.knife_lifecycle = textBox19.Text;
            newcode.knife_price = textBox18.Text;
            newcode.packing_type = comboBox3.SelectedIndex.ToString();
            newcode.packing_qty = textBox21.Text;
            newcode.rpm = textBox22.Text;
            newcode.pin_distance = textBox23.Text;
            newcode.process_type = textBox24.Text;
            newcode.eq1 = comboBox5.Text;
            newcode.eq2 = comboBox6.Text;
            newcode.steps = textBox30.Text;
            newcode.print_times = textBox29.Text;
            newcode.print_yn = checkBox2.Checked.ToString();
            newcode.draw_path = textBox1.Text;
            newcode.use_yn = checkBox1.Checked.ToString();
            newcode.remark = richTextBox1.Text;
            newcode.ins_empl = EMPL_NO;
            newcode.old_g_code = label38.Text;
            newcode.bom = dataGridView1;
            try
            {
                ProductBLL pro = new ProductBLL();
                if ((pro.checkBOM2(newcode.old_g_code, newcode.old_g_code.Substring(7, 1)) != 0) && !(newcode.old_g_code.Length < 8) && newcode.checkBOM()=="")
                {
                    try
                    {
                        newcode.updateBOM2();
                        pro.insertMaterialfromBOMtoMTable();
                        MessageBox.Show("Update BOM tính giá hoàn thành cho code:" + newcode.old_g_code);
                        //richTextBox1.Text = newcode.remark;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Lỗi M140, BOM2: " + ex.ToString());
                    }
                }
                else
                {
                    MessageBox.Show("Không có bom or chưa chọn code trong list or BOM chưa điền đủ thông tin");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            createNewCode();
        }

        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            //MessageBox.Show("Cell Enter");
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            richTextBox1.Text = dataGridView1.CurrentCell.Value.ToString();
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

        }
        private void button3_Click(object sender, EventArgs e)
        {

            string matsize = comboBox9.Text;
            string[] mat = matsize.Split('|');
            //this.dataGridView1.Rows.Add(mat[1], "", mat[0], "", "", "", mat[2], "", "", "");
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dt = pro.getMATInfo(mat[0]);
            string ssprice = "", cmsprice = "", slittingprice = "", masterwidth = "", roll_length = "", vendor = "";
            if (dt.Rows.Count > 0)
            {
                ssprice = dt.Rows[0]["SSPRICE"].ToString();
                cmsprice = dt.Rows[0]["CMSPRICE"].ToString();
                slittingprice = dt.Rows[0]["SLITTING_PRICE"].ToString();
                masterwidth = dt.Rows[0]["MASTER_WIDTH"].ToString();
                roll_length = dt.Rows[0]["ROLL_LENGTH"].ToString();
                vendor = dt.Rows[0]["CUST_CD"].ToString();
            }

            if (radioButton1.Checked == true)
            {
                this.dataGridView1.Rows.Add(mat[2], "LIEU", mat[0], "", mat[1], "", vendor, "", cmsprice, ssprice, slittingprice, masterwidth, roll_length, "1", "", "");
            }
            else if (radioButton2.Checked == true)
            {
                this.dataGridView1.Rows.Add(mat[2], "MUC", mat[0], "", mat[1], "", vendor, "", cmsprice, ssprice, slittingprice, masterwidth, roll_length, "1", "", "");
            }
            else
            {
                this.dataGridView1.Rows.Add(mat[2], "CORE", mat[0], "", mat[1], "", vendor, "", cmsprice, ssprice, slittingprice, masterwidth, roll_length, "1", "", "");
            }

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                ProductBLL pro = new ProductBLL();
                DataTable dt = pro.getMcodeInfo(textBox4.Text);
                dataGridView2.DataSource = dt;

            }
        }

        private void comboBox8_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {


        }
        ProductBLL pro = new ProductBLL();
        private void dataGridView2_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewRow row = dataGridView2.Rows[e.RowIndex];
            label38.Text = row.Cells["G_CODE"].Value.ToString();
            label39.Text = row.Cells["G_NAME"].Value.ToString();
            comboBox7.SelectedValue = row.Cells["CUST_CD"].Value.ToString();
            comboBox8.SelectedIndex = (int.Parse(row.Cells["CODE_12"].Value.ToString()) - 6);
            comboBox1.SelectedItem = row.Cells["PROD_TYPE"].Value.ToString();
            textBox2.Text = row.Cells["PROD_PROJECT"].Value.ToString();
            textBox3.Text = row.Cells["PROD_MODEL"].Value.ToString();
            textBox5.Text = row.Cells["G_NAME_KD"].Value.ToString();
            textBox6.Text = row.Cells["DESCR"].Value.ToString();
            textBox7.Text = row.Cells["PROD_MAIN_MATERIAL"].Value.ToString();
            textBox8.Text = row.Cells["G_NAME"].Value.ToString();
            textBox1.Text = row.Cells["DRAW_LINK"].Value.ToString();
            comboBox5.SelectedItem = (row.Cells["EQ1"].Value.ToString() == "" || row.Cells["EQ1"].Value.ToString() == "NO") ? "NO" : row.Cells["EQ1"].Value.ToString();
            comboBox6.SelectedItem = (row.Cells["EQ2"].Value.ToString() == "" || row.Cells["EQ2"].Value.ToString() == "NO") ? "NO" : row.Cells["EQ2"].Value.ToString();
            textBox30.Text = row.Cells["PROD_DIECUT_STEP"].Value.ToString();
            textBox29.Text = row.Cells["PROD_PRINT_TIMES"].Value.ToString();
            textBox16.Text = row.Cells["G_LENGTH"].Value.ToString();
            textBox15.Text = row.Cells["G_WIDTH"].Value.ToString();
            textBox14.Text = row.Cells["PD"].Value.ToString();
            textBox13.Text = row.Cells["G_C_R"].Value.ToString();// cavity hang
            textBox12.Text = row.Cells["G_C"].Value.ToString();// cavity cot
            textBox11.Text = row.Cells["G_LG"].Value.ToString();// line gap - khoang cach hang
            textBox10.Text = row.Cells["G_CG"].Value.ToString(); // column gap - khoang cach cot
            textBox9.Text = row.Cells["G_SG_L"].Value.ToString(); // side left gap - khoang cach toi liner trai
            textBox17.Text = row.Cells["G_SG_R"].Value.ToString(); // side right gap - khoang cach toi liner phai
            comboBox4.SelectedItem = (row.Cells["KNIFE_TYPE"].Value.ToString() == "" ? "NO" : row.Cells["KNIFE_TYPE"].Value.ToString()== "0" ? "PVC": row.Cells["KNIFE_TYPE"].Value.ToString() == "1" ? "PINACLE" : "NO"); // 
            textBox19.Text = row.Cells["KNIFE_LIFECYCLE"].Value.ToString();// tuoi dao
            textBox18.Text = row.Cells["KNIFE_PRICE"].Value.ToString();
            switch (row.Cells["CODE_33"].Value.ToString())
            {
                case "02":
                    comboBox3.SelectedItem = "ROLL";
                    break;
                case "03":
                    comboBox3.SelectedItem = "SHEET";
                    break;
                case "06":
                    comboBox3.SelectedItem = "PACK(BAG)";
                    break;
                case "04":
                    comboBox3.SelectedItem = "MET";
                    break;
                case "05":
                    comboBox3.SelectedItem = "BOX";
                    break;
                case "01":
                    comboBox3.SelectedItem = "EA";
                    break;
                case "07":
                    comboBox3.SelectedItem = "KG";
                    break;
                case "99":
                    comboBox3.SelectedItem = "X";
                    break;
                default:
                    comboBox3.SelectedItem = "X";
                    break;
            }

            textBox21.Text = row.Cells["ROLE_EA_QTY"].Value.ToString();
            textBox22.Text = row.Cells["RPM"].Value.ToString();
            textBox23.Text = row.Cells["PIN_DISTANCE"].Value.ToString();
            textBox24.Text = row.Cells["PROCESS_TYPE"].Value.ToString();
            checkBox2.Checked = (row.Cells["PRT_YN"].Value.ToString() == "Y" ? true : false);
            checkBox1.Checked = (row.Cells["USE_YN"].Value.ToString() == "Y" ? true : false);
            richTextBox1.Text = row.Cells["REMK"].Value.ToString();
            DataTable dt = new DataTable();
            dt = pro.getFullBOM(row.Cells["G_CODE"].Value.ToString(), row.Cells["REV_NO"].Value.ToString());
            if(checkBox3.Checked == true)
            {

            }
            else
            {
                dataGridView3.DataSource = dt;
            }
            

            dataGridView1.DataSource = null;
            /*xoa dong datagridview3 */
            if (dataGridView1.Rows.Count > 1)
            {
                do
                {
                    dataGridView1.Rows.RemoveAt(0);
                }
                while (dataGridView1.Rows.Count > 1);
            }

            dt = pro.getFullBOM2(row.Cells["G_CODE"].Value.ToString(), row.Cells["REV_NO"].Value.ToString());

            /*Them tung dong vao datagridview3*/

            if (dt.Rows.Count > 0)
            {

                List<string> list = new List<string>();
                list.Add("Hi");
                String[] str = list.ToArray();
              

                for (int kk = 00; kk < dt.Rows.Count; kk++)
                {
                    // DataRow drrow = bom2.NewRow();                    
                    List<string> list1 = new List<string>();
                    list1.Add(dt.Rows[kk]["M_CODE"].ToString());
                    list1.Add((dt.Rows[kk]["CATEGORY"].ToString() == "1" ? "LIEU" : dt.Rows[kk]["CATEGORY"].ToString() == "2" ? "MUC" : "CORE"));
                    list1.Add(dt.Rows[kk]["M_NAME"].ToString());
                    list1.Add(dt.Rows[kk]["USAGE"].ToString());
                    list1.Add(dt.Rows[kk]["MAT_CUTWIDTH"].ToString());
                    list1.Add(dt.Rows[kk]["MAT_THICKNESS"].ToString());
                    list1.Add(dt.Rows[kk]["CUST_CD"].ToString());
                    list1.Add(dt.Rows[kk]["IMPORT_CAT"].ToString());
                    list1.Add(dt.Rows[kk]["M_CMS_PRICE"].ToString());
                    list1.Add(dt.Rows[kk]["M_SS_PRICE"].ToString());
                    list1.Add(dt.Rows[kk]["M_SLITTING_PRICE"].ToString());
                    list1.Add(dt.Rows[kk]["MAT_MASTER_WIDTH"].ToString());
                    list1.Add(dt.Rows[kk]["MAT_ROLL_LENGTH"].ToString());
                    list1.Add(dt.Rows[kk]["M_QTY"].ToString());
                    list1.Add(dt.Rows[kk]["PROCESS_ORDER"].ToString());
                    list1.Add(dt.Rows[kk]["REMARK"].ToString());

                    String[] str1 = list1.ToArray();                  
                    dataGridView1.Rows.Add(str1);
                }                
            }
        }

        public bool active = true;
        private void button1_Click(object sender, EventArgs e)
        {
            active = !active;
            changeInputStatus(active);           
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (label38.Text.Length == 8)
            {
                addNewVer();
            }
            else
            {
                MessageBox.Show("Chọn code để up ver, và nhập thông tin trước khi up ver");
            }

        }

        private void button4_Click(object sender, EventArgs e)
        {           
            if (label38.Text.Length == 8)
            {
                updateCODEANDBOM();
            }
            else
            {
                MessageBox.Show("Chọn code để update trước khi update");
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (label38.Text.Length == 8)
            {
                addBOM2ofOldCode();
            }
            else
            {
                MessageBox.Show("Chọn code để add bom");
            }
            
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (label38.Text.Length == 8)
            {
                updateBOM2ofOldCode();
            }
            else
            {
                MessageBox.Show("Chọn code để add bom");
            }


        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {             
                System.Diagnostics.Process.Start(textBox1.Text);               
                
            }
            catch(Exception ex)
            {
                MessageBox.Show("Link lỗi:   " + ex.ToString());
            }            
            
        }

        private void button9_Click(object sender, EventArgs e)
        {
            try
            {                
                string Dir = System.IO.Directory.GetCurrentDirectory();
                //MessageBox.Show(Dir);
                string file = Dir + "\\BOMtemplate2.xlsx";
                string savebompath = "";              
                
                using (var fbd = new FolderBrowserDialog())
                {
                    DialogResult result = fbd.ShowDialog();

                    if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                    {
                        savebompath = fbd.SelectedPath;

                        ProductBLL pro = new ProductBLL();
                        DataTable dt = new DataTable();

                        //dataGridView1.Rows[rowindex].Cells[columnindex].Value.ToString();

                        var selectedRows = dataGridView2.SelectedRows
                       .OfType<DataGridViewRow>()
                       .Where(row => !row.IsNewRow)
                       .ToArray();

                        foreach (var row in selectedRows)
                        {
                            string g_code = row.Cells["G_CODE"].Value.ToString();
                            //MessageBox.Show(g_code);
                           //MessageBox.Show(dt.Rows.Count.ToString());

                            dt = pro.getFullInfoBOM2(g_code);

                            if (file != "" && dt.Rows.Count >=1)
                            {
                                ExcelFactory.editFileBOMExcel(file, dt, savebompath, dt.Rows[0]["M100_DRAW_LINK"].ToString());
                            }
                            else
                            {
                                MessageBox.Show("Chưa có BOM222 giá nên không xuất file được!");
                            }
                        }

                        MessageBox.Show("Export BOM hoàn thành !");                       
                        // MessageBox.Show(saveycsxpath);
                    }
                }



            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.ToString());
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            justCheckBOM();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            cloneBOM();
        }

        private void button12_Click(object sender, EventArgs e)
        {

        }
    }
    public class CMSCODE
    {
        public string cust_cd { get; set; }
        public string g_code { get; set; }
        public string project { get; set; }
        public string model { get; set; }

        public string prod_type { get; set; }
        public string dactinh { get; set; }
        public string phanloai { get; set; }
        public string g_name { get; set; }
        public string g_name_kd { get; set; }
        public string description { get; set; }
        public string main_material { get; set; }
        public string width { get; set; }
        public string length { get; set; }
        public string feeding { get; set; }
        public string cavity_hang { get; set; }
        public string cavity_cot { get; set; }
        public string k_c_hang { get; set; }
        public string k_c_cot { get; set; }
        public string k_c_liner_trai { get; set; }
        public string k_c_liner_phai { get; set; }
        public string knife { get; set; }
        public string knife_lifecycle { get; set; }
        public string knife_price { get; set; }
        public string packing_type { get; set; }
        public string packing_qty { get; set; }
        public string rpm { get; set; }
        public string pin_distance { get; set; }
        public string process_type { get; set; }
        public string roll_open_direction { get; set; }
        public string eq1 { get; set; }
        public string eq2 { get; set; }

        public string steps { get; set; }
        public string print_yn { get; set; }
        public string print_times { get; set; }

        public string draw_path { get; set; }
        public string use_yn { get; set; }
        public string remark { get; set; }

        public string ins_empl { get; set; }

        public string old_g_code { get; set; }

        public DataGridView bom { get; set; }

       

        public string checkBOM()
        {
            string checkBOMOK = "";

            if (this.bom.Rows.Count > 1)
            {
                int seq = 1;
                ProductBLL pro = new ProductBLL();

                foreach (DataGridViewRow row in this.bom.Rows)
                {
                    if (!row.IsNewRow)
                    {
                        string
                            CATEGORY = row.Cells["CATEGORY"].Value.ToString(),
                            M_CODE = row.Cells["MATCODE"].Value.ToString(),
                            M_NAME = row.Cells["MATNAME"].Value.ToString(),
                            CUST_CD = row.Cells["VENDOR"].Value.ToString(),
                            USAGE = row.Cells["USAGE"].Value.ToString(),
                            CUTWIDTH = row.Cells["CUTWIDTH"].Value.ToString(),
                            MAT_THICKNESS = row.Cells["THICKNESS"].Value.ToString(),
                            M_QTY = "1",
                            PROCESS_ORDER = row.Cells["PROCESS_ORDER"].Value.ToString();
                           
                        if(CATEGORY == "")
                        {
                            row.Cells["CATEGORY"].Style.BackColor = Color.Red;
                            checkBOMOK = "Thiếu thông tin trong BOM, hãy điền tất cả";
                        }
                        if (M_CODE == "")
                        {
                            row.Cells["MATCODE"].Style.BackColor = Color.Red;
                            checkBOMOK = "Thiếu thông tin trong BOM, hãy điền tất cả";
                        }
                        if (M_NAME == "")
                        {
                            row.Cells["MATNAME"].Style.BackColor = Color.Red;
                            checkBOMOK = "Thiếu thông tin trong BOM, hãy điền tất cả";
                        }
                        if (CUST_CD == "")
                        {
                            row.Cells["VENDOR"].Style.BackColor = Color.Red;
                            checkBOMOK = "Thiếu thông tin trong BOM, hãy điền tất cả";
                        }
                        if (USAGE == "")
                        {
                            row.Cells["USAGE"].Style.BackColor = Color.Red;
                            checkBOMOK = "Thiếu thông tin trong BOM, hãy điền tất cả";
                        }
                        if (CUTWIDTH == "")
                        {
                            row.Cells["CUTWIDTH"].Style.BackColor = Color.Red;
                            checkBOMOK = "Thiếu thông tin trong BOM, hãy điền tất cả";
                        }
                        if (MAT_THICKNESS == "")
                        {
                            row.Cells["THICKNESS"].Style.BackColor = Color.Red;
                            checkBOMOK = "Thiếu thông tin trong BOM, hãy điền tất cả";
                        }
                        if (PROCESS_ORDER == "")
                        {
                            row.Cells["PROCESS_ORDER"].Style.BackColor = Color.Red;
                            checkBOMOK = "Thiếu thông tin trong BOM, hãy điền tất cả";
                        }                      

                    }
                }
            }
            else
            {
                MessageBox.Show("Không có dòng nào được thêm");
            }
            return checkBOMOK;


        }
        public string checkInput()
        {
            string inputOK = "";
            if (cust_cd == "") inputOK += "cust_cd không được để trống_";
            if (g_code == "") inputOK += "g_code không được để trống_";
            if (project == "") inputOK += "project không được để trống_";
            if (model == "") inputOK += "model không được để trống_";
            if (dactinh == "") inputOK += "dactinh không được để trống_";
            if (phanloai == "") inputOK += "phanloai không được để trống_";
            if (g_name == "") inputOK += "g_name không được để trống_";
            if (g_name_kd == "") inputOK += "g_name_kd không được để trống_";
            if (description == "") inputOK += "description không được để trống_";
            if (main_material == "") inputOK += "main_material không được để trống_";
            if (width == "") inputOK += "width không được để trống_";
            if (length == "") inputOK += "length không được để trống_";
            if (feeding == "") inputOK += "feeding không được để trống_";
            if (cavity_hang == "") inputOK += "cavity_hang không được để trống_";
            if (cavity_cot == "") inputOK += "cavity_cot không được để trống_";
            if (k_c_hang == "") inputOK += "k_c_hang không được để trống_";
            if (k_c_cot == "") inputOK += "k_c_cot không được để trống_";
            if (k_c_liner_trai == "") inputOK += "k_c_liner_trai không được để trống_";
            if (k_c_liner_phai == "") inputOK += "k_c_liner_phai không được để trống_";
            if (knife == "") inputOK += "knife không được để trống_";
            if (knife_lifecycle == "") inputOK += "knife_lifecycle không được để trống_";
            if (knife_price == "") inputOK += "knife_price không được để trống_";
            if (packing_type == "") inputOK += "packing_type không được để trống_";
            if (packing_qty == "") inputOK += "packing_qty không được để trống_";
            if (rpm == "") inputOK += "rpm không được để trống_";
            if (pin_distance == "") inputOK += "pin_distance không được để trống_";
            if (process_type == "") inputOK += "process_type không được để trống_";
            if (eq1 == "") inputOK += "eq1 không được để trống_";
            if (eq2 == "") inputOK += "eq2 không được để trống_";
            if (steps == "") inputOK += "steps không được để trống_";
            if (print_yn == "") inputOK += "print_yn không được để trống_";
            if (print_times == "") inputOK += "print_times không được để trống_";
            if (draw_path == "") inputOK += "draw_path không được để trống_";
            if (use_yn == "") inputOK += "use_yn không được để trống_";
            if (ins_empl == "") inputOK += "ins_empl không được để trống_";
            return inputOK;
        }
        public bool insertVerCMS()
        {
            bool result = true;

            ProductBLL pro = new ProductBLL();
            string new_g_code = "", dactinh = "", phanloai = "", packingtype = "", print_yn = "", use_yn = "";

            switch (this.dactinh)
            {
                case "0": //thanh pham
                    dactinh = "6";
                    break;
                case "1": // ban thanh pham
                    dactinh = "7";
                    break;
                case "2": // nguyen chiec k fai ribbon
                    dactinh = "8";
                    break;
                case "3": // nguyen chiec ribbon
                    dactinh = "9";
                    break;
                default:

                    break;
            }


            switch (this.phanloai)
            {
                case "0":  //tsp
                    phanloai = "C";
                    break;
                case "1": //label
                    phanloai = "A";
                    break;
                case "2": //uv
                    phanloai = "C";
                    break;
                case "3": // oled
                    phanloai = "C";
                    break;
                case "4": //tape
                    phanloai = "B";
                    break;
                case "5":
                    phanloai = "E";
                    break;
                case "6":
                    phanloai = "Z";
                    break;
                default:

                    break;
            }
            switch (this.packing_type)
            {
                case "0":  //roll
                    packingtype = "02";
                    break;
                case "1": //sheet
                    packingtype = "03";
                    break;
                case "2": //PACK(BAG)
                    packingtype = "06";
                    break;
                case "3": // MET
                    packingtype = "04";
                    break;
                case "4": //BOX
                    packingtype = "05";
                    break;
                case "5": //EA
                    packingtype = "01";
                    break;
                case "6": //KG
                    packingtype = "07";
                    break;
                case "7": //X
                    packingtype = "99";
                    break;
                default:

                    break;
            }
            switch (this.print_yn)
            {
                case "True":
                    print_yn = "Y";
                    break;

                case "False":
                    print_yn = "N";
                    break;

            }
            switch (this.use_yn)
            {
                case "True":
                    use_yn = "Y";
                    break;

                case "False":
                    use_yn = "N";
                    break;

            }


            string verArray = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            string lastVer = pro.getlastver(this.old_g_code.Substring(0, 7));
            int lastverPos = verArray.IndexOf(lastVer[0]);
            string nextVer = verArray.Substring(lastverPos + 1, 1);
            string last_seq_no = pro.getLastG_CODE_SEQ_NO(dactinh, phanloai);

            string next_seq_no = "";

            if (dactinh == "9")
            {
                next_seq_no = (int.Parse(last_seq_no) + 1).ToString("000000");
                new_g_code = dactinh + phanloai + next_seq_no;
            }
            else
            {
                next_seq_no = this.old_g_code.Substring(2, 5);
                new_g_code = this.old_g_code.Substring(0, 7) + nextVer;
            }

            this.g_code = new_g_code;

            if (this.g_name == "" || this.g_name_kd == "" || this.model == "" || this.description == "")
            {
                MessageBox.Show("Phải điền tất cả các trường được bắt buộc !");
            }
            else
            {
                string m100_values = $"('002','{new_g_code}','{this.g_name}','{dactinh}','{next_seq_no}','{nextVer}','{packingtype}','{this.cust_cd}','{this.g_name}','','','{phanloai}','AA','','{print_yn}','{this.print_times}','',{this.packing_qty},{this.width},{this.length},0.0,{this.cavity_cot},{this.k_c_hang},{this.k_c_liner_trai},{this.k_c_liner_phai},{this.k_c_cot},'{this.remark}','{use_yn}','{this.ins_empl}', GETDATE(),'{this.ins_empl}',GETDATE(),'{this.project}','{this.model}','{this.g_name_kd}','{this.draw_path}','{this.eq1}','{this.eq2}','{this.steps}', '{this.feeding}','{this.knife}','{this.knife_lifecycle}','{this.knife_price}','{this.rpm}','{this.pin_distance}','{this.process_type}','{this.cavity_hang}','{this.description}','{this.main_material}','{this.prod_type}')";
                //MessageBox.Show(m100_values);
                this.remark = m100_values;
                try
                {
                    pro.M100_insert(m100_values);
                    result = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi M100: " + ex.ToString());
                    result = false;
                }

            }
            return result;
        }

        public bool updateCode()
        {
            bool result = true;

            ProductBLL pro = new ProductBLL();
            string new_g_code = "", dactinh = "", phanloai = "", packingtype = "", print_yn = "", use_yn = "";

            switch (this.dactinh)
            {
                case "0": //thanh pham
                    dactinh = "6";
                    break;
                case "1": // ban thanh pham
                    dactinh = "7";
                    break;
                case "2": // nguyen chiec k fai ribbon
                    dactinh = "8";
                    break;
                case "3": // nguyen chiec ribbon
                    dactinh = "9";
                    break;
                default:

                    break;
            }


            switch (this.phanloai)
            {
                case "0":  //tsp
                    phanloai = "C";
                    break;
                case "1": //label
                    phanloai = "A";
                    break;
                case "2": //uv
                    phanloai = "C";
                    break;
                case "3": // oled
                    phanloai = "C";
                    break;
                case "4": //tape
                    phanloai = "B";
                    break;
                case "5":
                    phanloai = "E";
                    break;
                case "6":
                    phanloai = "Z";
                    break;
                default:

                    break;
            }
            switch (this.packing_type)
            {
                case "0":  //roll
                    packingtype = "02";
                    break;
                case "1": //sheet
                    packingtype = "03";
                    break;
                case "2": //PACK(BAG)
                    packingtype = "06";
                    break;
                case "3": // MET
                    packingtype = "04";
                    break;
                case "4": //BOX
                    packingtype = "05";
                    break;
                case "5": //EA
                    packingtype = "01";
                    break;
                case "6": //KG
                    packingtype = "07";
                    break;
                case "7": //X
                    packingtype = "99";
                    break;
                default:

                    break;
            }
            switch (this.print_yn)
            {
                case "True":
                    print_yn = "Y";
                    break;

                case "False":
                    print_yn = "N";
                    break;

            }
            switch (this.use_yn)
            {
                case "True":
                    use_yn = "Y";
                    break;

                case "False":
                    use_yn = "N";
                    break;

            }

            this.g_code = this.old_g_code;


            string next_seq_no = "", nextVer = "";

            if (dactinh == "9")
            {
                next_seq_no = this.old_g_code.Substring(2, 6);                
               
            }
            else
            {
                next_seq_no = this.old_g_code.Substring(2, 5);
                nextVer = this.old_g_code.Substring(7, 1);
            }
           
  string m100_values = $" CTR_CD='002',G_CODE='{this.old_g_code}',G_NAME='{this.g_name}',CODE_12='{dactinh}',SEQ_NO='{next_seq_no}',REV_NO='{nextVer}',CODE_33='{packingtype}',CUST_CD='{this.cust_cd}',G_CODE_C='{this.g_name}',G_CODE_V='',G_CODE_K='',CODE_27='{phanloai}',CODE_28='AA',PRT_DRT='',PRT_YN='{print_yn}',PROD_PRINT_TIMES='{this.print_times}',PACK_DRT='',ROLE_EA_QTY={this.packing_qty},G_WIDTH={this.width},G_LENGTH={this.length},G_R=0.0,G_C={this.cavity_cot},G_LG={this.k_c_hang},G_SG_L={this.k_c_liner_trai},G_SG_R={this.k_c_liner_phai},G_CG={this.k_c_cot},REMK='{this.remark}',USE_YN='{use_yn}',INS_EMPL='{this.ins_empl}', INS_DATE= GETDATE(),UPD_EMPL='{this.ins_empl}',UPD_DATE=GETDATE(),PROD_PROJECT='{this.project}',PROD_MODEL='{this.model}',G_NAME_KD='{this.g_name_kd}',DRAW_LINK='{this.draw_path}',EQ1='{this.eq1}',EQ2='{this.eq2}',PROD_DIECUT_STEP='{this.steps}',PD= '{this.feeding}',KNIFE_TYPE='{this.knife}',KNIFE_LIFECYCLE='{this.knife_lifecycle}',KNIFE_PRICE='{this.knife_price}',RPM='{this.rpm}',PIN_DISTANCE='{this.pin_distance}',PROCESS_TYPE='{this.process_type}',G_C_R='{this.cavity_hang}',DESCR='{this.description}',PROD_MAIN_MATERIAL='{this.main_material}', PROD_TYPE='{this.prod_type}' WHERE G_CODE = '{this.old_g_code}'";
                //MessageBox.Show(m100_values);
                this.remark = m100_values;
                try
                {
                    pro.M100_update(m100_values);
                    result = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi ngoại lệ: " + ex.ToString());
                    result = false;
                }

            
            return result;
        }


        public bool insertOldCMS()
        {
            bool result = true;

            ProductBLL pro = new ProductBLL();
            string new_g_code = "", dactinh = "", phanloai = "", packingtype = "", print_yn = "", use_yn = "";

            switch (this.dactinh)
            {
                case "0": //thanh pham
                    dactinh = "6";
                    break;
                case "1": // ban thanh pham
                    dactinh = "7";
                    break;
                case "2": // nguyen chiec k fai ribbon
                    dactinh = "8";
                    break;
                case "3": // nguyen chiec ribbon
                    dactinh = "9";
                    break;
                default:

                    break;
            }


            switch (this.phanloai)
            {
                case "0":  //tsp
                    phanloai = "C";
                    break;
                case "1": //label
                    phanloai = "A";
                    break;
                case "2": //uv
                    phanloai = "C";
                    break;
                case "3": // oled
                    phanloai = "C";
                    break;
                case "4": //tape
                    phanloai = "B";
                    break;
                case "5":
                    phanloai = "E";
                    break;
                case "6":
                    phanloai = "Z";
                    break;
                default:

                    break;
            }
            switch (this.packing_type)
            {
                case "0":  //roll
                    packingtype = "02";
                    break;
                case "1": //sheet
                    packingtype = "03";
                    break;
                case "2": //PACK(BAG)
                    packingtype = "06";
                    break;
                case "3": // MET
                    packingtype = "04";
                    break;
                case "4": //BOX
                    packingtype = "05";
                    break;
                case "5": //EA
                    packingtype = "01";
                    break;
                case "6": //KG
                    packingtype = "07";
                    break;
                case "7": //X
                    packingtype = "99";
                    break;
                default:

                    break;
            }
            switch (this.print_yn)
            {
                case "True":
                    print_yn = "Y";
                    break;

                case "False":
                    print_yn = "N";
                    break;

            }
            switch (this.use_yn)
            {
                case "True":
                    use_yn = "Y";
                    break;

                case "False":
                    use_yn = "N";
                    break;

            }




            string last_seq_no = pro.getLastG_CODE_SEQ_NO(dactinh, phanloai);
            string next_seq_no = "";

            if (last_seq_no != "")
            {

                if (dactinh == "9")
                {
                    next_seq_no = (int.Parse(last_seq_no) + 1).ToString("000000");
                    new_g_code = dactinh + phanloai + next_seq_no;
                }
                else
                {
                    next_seq_no = (int.Parse(last_seq_no) + 1).ToString("00000");
                    new_g_code = dactinh + phanloai + next_seq_no + "A";
                }

            }
            else
            {
                if (dactinh == "9")
                {
                    next_seq_no = "000001";
                    new_g_code = dactinh + phanloai + next_seq_no;
                }
                else
                {
                    next_seq_no = "00001" + "A";
                    new_g_code = dactinh + phanloai + next_seq_no + "A";
                }
            }

            this.g_code = new_g_code;

            if (2 == 0)
            {
                MessageBox.Show("Phải điền tất cả các trường được bắt buộc !");
            }
            else
            {
                string m100_values = $"('002','{new_g_code}','{this.g_name}','{dactinh}','{next_seq_no}','A','{packingtype}','{this.cust_cd}','{this.g_name}','','','{phanloai}','AA','','{print_yn}','{this.print_times}','',{this.packing_qty},{this.width},{this.length},0.0,{this.cavity_cot},{this.k_c_hang},{this.k_c_liner_trai},{this.k_c_liner_phai},{this.k_c_cot},'{this.remark}','{use_yn}','{this.ins_empl}', GETDATE(),'{this.ins_empl}',GETDATE(),'{this.project}','{this.model}','{this.g_name_kd}','{this.draw_path}','{this.eq1}','{this.eq2}','{this.steps}', '{this.feeding}','{this.knife}','{this.knife_lifecycle}','{this.knife_price}','{this.rpm}','{this.pin_distance}','{this.process_type}','{this.cavity_hang}','{this.description}','{this.main_material}', '{this.prod_type}')";
                //MessageBox.Show(m100_values);
                this.remark = m100_values;
                try
                {
                    pro.M100_insert(m100_values);
                    result = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi M100: " + ex.ToString());
                    result = false;
                }



            }
            return result;
        }
        private string padding_num(int number)
        {
            string output = "";
            if (number < 10)
            {
                return "00" + number;
            }
            else if (number < 100)
            {
                return "0" + number;
            }
            else if (number < 1000)
            {
                return "" + number;
            }
            return output;
        }
        public void insertOldBOM()
        {
            if (this.bom.Rows.Count > 1)
            {
                int seq = 1;
                ProductBLL pro = new ProductBLL();

                foreach (DataGridViewRow row in this.bom.Rows)
                {
                    if (!row.IsNewRow)
                    {
                        string
                            G_CODE = this.g_code,
                            RIV_NO = this.g_code.Substring(7,1),
                            G_SEQ = padding_num(seq),
                            M_CODE = row.Cells["MATCODE"].Value.ToString(),
                            M_QTY = "1",
                            META_PAT_CD = "X",
                            REMK = row.Cells["REMARK"].Value.ToString(),
                            USE_YN = "Y",
                            INS_EMPL = this.ins_empl,
                            UPD_EMPL = this.ins_empl;
                        string insertBOMVALUE = $"('002','{G_CODE}','{RIV_NO}','{G_SEQ}','{M_CODE}','{M_QTY}','{META_PAT_CD}', '{REMK}','{USE_YN}',GETDATE(),'{INS_EMPL}',GETDATE(),'{UPD_EMPL}')";
                        pro.insertOldBOM(insertBOMVALUE);
                        seq++;
                    }
                }
            }
            else
            {
                MessageBox.Show("Đã thêm code mà không thêm BOM");
            }
        }

        public void insertNewBOM()
        {
            if (this.bom.Rows.Count > 1)
            {
                int seq = 1;
                ProductBLL pro = new ProductBLL();

                foreach (DataGridViewRow row in this.bom.Rows)
                {
                    if (!row.IsNewRow)
                    {
                        string
                            G_CODE = this.g_code,
                            RIV_NO = this.g_code.Substring(7,1),
                            G_SEQ = padding_num(seq),
                            CATEGORY = row.Cells["CATEGORY"].Value.ToString(),
                            M_CODE = row.Cells["MATCODE"].Value.ToString(),
                            M_NAME = row.Cells["MATNAME"].Value.ToString(),
                            CUST_CD = row.Cells["VENDOR"].Value.ToString(),
                            CMS_PRICE = row.Cells["CMS_PRICE"].Value.ToString(),
                            SS_PRICE = row.Cells["SS_PRICE"].Value.ToString(),
                            SLITTING_PRICE = row.Cells["SLITTING_PRICE"].Value.ToString(),
                            USAGE = row.Cells["USAGE"].Value.ToString(),
                            MASTERWIDTH = row.Cells["MASTER_WIDTH"].Value.ToString(),
                            CUTWIDTH = row.Cells["CUTWIDTH"].Value.ToString(),
                            ROLL_LENGTH = row.Cells["ROLL_LENGTH"].Value.ToString(),
                            MAT_THICKNESS = row.Cells["THICKNESS"].Value.ToString(),
                            M_QTY = row.Cells["M_QTY"].Value.ToString(),
                            REMK = row.Cells["REMARK"].Value.ToString(),
                            PROCESS_ORDER = row.Cells["PROCESS_ORDER"].Value.ToString(),
                            INS_EMPL = this.ins_empl,
                            UPD_EMPL = this.ins_empl;
                        int categorynum = 1;

                        if (CATEGORY == "LIEU")
                        {
                            categorynum = 1;
                        }
                        else if (CATEGORY == "MUC")
                        {
                            categorynum = 2;
                        }
                        else if (CATEGORY == "CORE")
                        {
                            categorynum = 3;
                        }
                        else
                        {
                            categorynum = 1;
                        }

                        string insertBOMVALUE = $"('002','{G_CODE}','{RIV_NO}','{G_SEQ}','{categorynum}','{M_CODE}','{M_NAME}','{CUST_CD}','SK','{CMS_PRICE}','{SS_PRICE}','{SLITTING_PRICE}','{USAGE}','{MASTERWIDTH}','{CUTWIDTH}','{ROLL_LENGTH}','{MAT_THICKNESS}','{M_QTY}', '{REMK}','{PROCESS_ORDER}',GETDATE(),'{INS_EMPL}',GETDATE(),'{UPD_EMPL}')";

                        pro.insertNewBOM(insertBOMVALUE);
                        seq++;
                    }
                }
            }
            else
            {
                MessageBox.Show("Đã thêm code mà không thêm BOM tính giá");
            }
        }

        public void insertBOM2OldCode()
        {
            if (this.bom.Rows.Count > 1)
            {
                int seq = 1;
                ProductBLL pro = new ProductBLL();

                foreach (DataGridViewRow row in this.bom.Rows)
                {
                    if (!row.IsNewRow)
                    {
                        string
                            G_CODE = this.old_g_code,
                            RIV_NO = this.old_g_code.Substring(7,1),
                            G_SEQ = padding_num(seq),
                            CATEGORY = row.Cells["CATEGORY"].Value.ToString(),
                            M_CODE = row.Cells["MATCODE"].Value.ToString(),
                            M_NAME = row.Cells["MATNAME"].Value.ToString(),
                            CUST_CD = row.Cells["VENDOR"].Value.ToString(),
                            CMS_PRICE = row.Cells["CMS_PRICE"].Value.ToString(),
                            SS_PRICE = row.Cells["SS_PRICE"].Value.ToString(),
                            SLITTING_PRICE = row.Cells["SLITTING_PRICE"].Value.ToString(),
                            USAGE = row.Cells["USAGE"].Value.ToString(),
                            MASTERWIDTH = row.Cells["MASTER_WIDTH"].Value.ToString(),
                            CUTWIDTH = row.Cells["CUTWIDTH"].Value.ToString(),
                            ROLL_LENGTH = row.Cells["ROLL_LENGTH"].Value.ToString(),
                            MAT_THICKNESS = row.Cells["THICKNESS"].Value.ToString(),
                            M_QTY = row.Cells["M_QTY"].Value.ToString(),
                            REMK = row.Cells["REMARK"].Value.ToString(),
                            PROCESS_ORDER = row.Cells["PROCESS_ORDER"].Value.ToString(),
                            INS_EMPL = this.ins_empl,
                            UPD_EMPL = this.ins_empl;
                        int categorynum = 1;

                        if (CATEGORY == "LIEU")
                        {
                            categorynum = 1;
                        }
                        else if (CATEGORY == "MUC")
                        {
                            categorynum = 2;
                        }
                        else if (CATEGORY == "CORE")
                        {
                            categorynum = 3;
                        }
                        else
                        {
                            categorynum = 1;
                        }

                        string insertBOMVALUE = $"('002','{G_CODE}','{RIV_NO}','{G_SEQ}','{categorynum}','{M_CODE}','{M_NAME}','{CUST_CD}','SK','{CMS_PRICE}','{SS_PRICE}','{SLITTING_PRICE}','{USAGE}','{MASTERWIDTH}','{CUTWIDTH}','{ROLL_LENGTH}','{MAT_THICKNESS}','{M_QTY}', '{REMK}','{PROCESS_ORDER}',GETDATE(),'{INS_EMPL}',GETDATE(),'{UPD_EMPL}')";

                        pro.insertNewBOM(insertBOMVALUE);
                        seq++;
                    }
                }
            }
            else
            {
                MessageBox.Show("Không có dòng nào được thêm");
            }
        }

        public void updateBOM2()
        {
            if (this.bom.Rows.Count > 1)
            {
                int seq = 1;
                ProductBLL pro = new ProductBLL();
                pro.deleteBOM(this.old_g_code);

                foreach (DataGridViewRow row in this.bom.Rows)
                {
                    if (!row.IsNewRow)
                    {
                        string
                            G_CODE = this.old_g_code,
                            RIV_NO = this.old_g_code.Substring(7, 1),
                            G_SEQ = padding_num(seq),
                            CATEGORY = row.Cells["CATEGORY"].Value.ToString(),
                            M_CODE = row.Cells["MATCODE"].Value.ToString(),
                            M_NAME = row.Cells["MATNAME"].Value.ToString(),
                            CUST_CD = row.Cells["VENDOR"].Value.ToString(),
                            CMS_PRICE = row.Cells["CMS_PRICE"].Value.ToString(),
                            SS_PRICE = row.Cells["SS_PRICE"].Value.ToString(),
                            SLITTING_PRICE = row.Cells["SLITTING_PRICE"].Value.ToString(),
                            USAGE = row.Cells["USAGE"].Value.ToString(),
                            MASTERWIDTH = row.Cells["MASTER_WIDTH"].Value.ToString(),
                            CUTWIDTH = row.Cells["CUTWIDTH"].Value.ToString(),
                            ROLL_LENGTH = row.Cells["ROLL_LENGTH"].Value.ToString(),
                            MAT_THICKNESS = row.Cells["THICKNESS"].Value.ToString(),
                            M_QTY = row.Cells["M_QTY"].Value.ToString(),
                            REMK = row.Cells["REMARK"].Value.ToString(),
                            PROCESS_ORDER = row.Cells["PROCESS_ORDER"].Value.ToString(),
                            INS_EMPL = this.ins_empl,
                            UPD_EMPL = this.ins_empl;
                        int categorynum = 1;

                        if (CATEGORY == "LIEU")
                        {
                            categorynum = 1;
                        }
                        else if (CATEGORY == "MUC")
                        {
                            categorynum = 2;
                        }
                        else if (CATEGORY == "CORE")
                        {
                            categorynum = 3;
                        }
                        else
                        {
                            categorynum = 1;
                        }

                        string insertBOMVALUE = $"('002','{G_CODE}','{RIV_NO}','{G_SEQ}','{categorynum}','{M_CODE}','{M_NAME}','{CUST_CD}','SK','{CMS_PRICE}','{SS_PRICE}','{SLITTING_PRICE}','{USAGE}','{MASTERWIDTH}','{CUTWIDTH}','{ROLL_LENGTH}','{MAT_THICKNESS}','{M_QTY}', '{REMK}','{PROCESS_ORDER}',GETDATE(),'{INS_EMPL}',GETDATE(),'{UPD_EMPL}')";
                       // this.remark = insertBOMVALUE;
                        pro.insertNewBOM(insertBOMVALUE);
                        seq++;
                    }
                }
            }
            else
            {
                MessageBox.Show("Đã xóa BOM tính giá");
            }
        }

    }
}
