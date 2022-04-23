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
    public partial class TinhBaoGia : Form
    {
        public TinhBaoGia()
        {
            InitializeComponent();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        //test commit
        public void loadcodelist()
        {
            ProductBLL pro = new ProductBLL();
            DataTable dt = pro.getcodebom2Info(textBox1.Text,(checkBox2.Checked ==true ? "chuatinhgia":"tinhgiaroi"));
            dataGridView1.DataSource = dt;
            formatcodelist(dataGridView1);
        }
        private void button1_Click(object sender, EventArgs e)
        {
            loadcodelist();
        }

        private void textBox1_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                loadcodelist();
            }
        }

        private void dataGridView1_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridViewRow row = dataGridView1.Rows[e.RowIndex];
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();
                dt = pro.getBOM2Info(row.Cells["G_CODE"].Value.ToString(), "lieu");
                dataGridView2.DataSource = dt;
                formatbomlist(dataGridView2);

                DataGridViewRow row1 = dataGridView1.Rows[e.RowIndex];
                label24.Text = "Code: " + row1.Cells["G_NAME"].Value.ToString();
                label26.Text = "Model: " + row1.Cells["PROD_MODEL"].Value.ToString();
                label25.Text = "Khách hàng: " + row1.Cells["CUST_NAME_KD"].Value.ToString();
                label27.Text = "Product Type: " + row1.Cells["PROD_TYPE"].Value.ToString();
                label28.Text = "Giá (CMS): " + row1.Cells["PRODUCT_CMSPRICE"].Value.ToString();
                label29.Text = "Giá (Samsung): " + row1.Cells["PRODUCT_SSPRICE"].Value.ToString();
                //MessageBox.Show(row1.Cells["PROD_TYPE"].Value.ToString());
                setconfig(row1.Cells["PROD_TYPE"].Value.ToString());                
            }
            catch(NullReferenceException ex)
            {
               
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.ToString());
            }
        }
        public void formatcodelist(DataGridView dataGridView1)
        {
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;

            dataGridView1.Columns["CUST_NAME_KD"].DefaultCellStyle.BackColor = Color.Aqua;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.BackColor = Color.Brown;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.ForeColor = Color.Yellow;


            dataGridView1.Columns["MATERIAL_COST_CMS"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["MATERIAL_COST_CMS"].DefaultCellStyle.ForeColor = Color.White;

            dataGridView1.Columns["PROCESS_COST_CMS"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["PROCESS_COST_CMS"].DefaultCellStyle.ForeColor = Color.White;

            dataGridView1.Columns["OTHER_COST_CMS"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["OTHER_COST_CMS"].DefaultCellStyle.ForeColor = Color.White;

            dataGridView1.Columns["PROFIT_VALUE_CMS"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["PROFIT_VALUE_CMS"].DefaultCellStyle.ForeColor = Color.White;

            dataGridView1.Columns["MCR_CMS"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["MCR_CMS"].DefaultCellStyle.ForeColor = Color.White;


            dataGridView1.Columns["MATERIAL_COST_SS"].DefaultCellStyle.BackColor = Color.DarkGray;
            dataGridView1.Columns["MATERIAL_COST_SS"].DefaultCellStyle.ForeColor = Color.White;

            dataGridView1.Columns["PROCESS_COST_SS"].DefaultCellStyle.BackColor = Color.DarkGray;
            dataGridView1.Columns["PROCESS_COST_SS"].DefaultCellStyle.ForeColor = Color.White;

            dataGridView1.Columns["OTHER_COST_SS"].DefaultCellStyle.BackColor = Color.DarkGray;
            dataGridView1.Columns["OTHER_COST_SS"].DefaultCellStyle.ForeColor = Color.White;

            dataGridView1.Columns["PROFIT_VALUE_SS"].DefaultCellStyle.BackColor = Color.DarkGray;
            dataGridView1.Columns["PROFIT_VALUE_SS"].DefaultCellStyle.ForeColor = Color.White;

            dataGridView1.Columns["MCR_SS"].DefaultCellStyle.BackColor = Color.DarkGray;
            dataGridView1.Columns["MCR_SS"].DefaultCellStyle.ForeColor = Color.White;

            dataGridView1.Columns["PRODUCT_CMSPRICE"].DefaultCellStyle.BackColor = Color.Orange;
            dataGridView1.Columns["PRODUCT_CMSPRICE"].DefaultCellStyle.ForeColor = Color.Black;

            dataGridView1.Columns["PRODUCT_SSPRICE"].DefaultCellStyle.BackColor = Color.LightBlue;
            dataGridView1.Columns["PRODUCT_SSPRICE"].DefaultCellStyle.ForeColor = Color.Black;

        }
        public void formatbomlist(DataGridView dataGridView1)
        {
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;

           
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.BackColor = Color.Brown;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.ForeColor = Color.Yellow;

            dataGridView1.Columns["M_NAME"].DefaultCellStyle.BackColor = Color.Brown;
            dataGridView1.Columns["M_NAME"].DefaultCellStyle.ForeColor = Color.LightGreen;

            dataGridView1.Columns["M_CMS_PRICE"].DefaultCellStyle.BackColor = Color.LightBlue;
            dataGridView1.Columns["M_CMS_PRICE"].DefaultCellStyle.ForeColor = Color.Black;

            dataGridView1.Columns["M_SS_PRICE"].DefaultCellStyle.BackColor = Color.LightBlue;
            dataGridView1.Columns["M_SS_PRICE"].DefaultCellStyle.ForeColor = Color.Black;

            dataGridView1.Columns["M_SLITTING_PRICE"].DefaultCellStyle.BackColor = Color.LightBlue;
            dataGridView1.Columns["M_SLITTING_PRICE"].DefaultCellStyle.ForeColor = Color.Black;

            dataGridView1.Columns["MAT_CUTWIDTH"].DefaultCellStyle.BackColor = Color.LightBlue;
            dataGridView1.Columns["MAT_CUTWIDTH"].DefaultCellStyle.ForeColor = Color.Black;


        }

        public void tinhgia(int option)
        {
            ProductBLL pro = new ProductBLL();

            float management_rate = 0.08f;
            float delivery_rate = 0.01f;
            float packing_rate = 0.0f;
            float inspection_rate = 0.02f;
            float profit_rate = 0.1f;
            float KDrate = 0.15f;
            float hesonhancong = 1.0f;


            profit_rate = float.Parse(textBox24.Text)/100; //profit rate
            delivery_rate = float.Parse(textBox25.Text)/100;    //delivery_rate     
            management_rate= float.Parse(textBox30.Text)/100; //management_rate
            KDrate = float.Parse(textBox35.Text)/100; //KD-rate
            packing_rate = float.Parse(textBox36.Text)/100; //packing_rate
            inspection_rate = float.Parse(textBox31.Text)/100; //inspection_rate
            hesonhancong = float.Parse(textBox27.Text); //he so nhan cong

            float material_cost_cms = 0.0f;
            float material_cost_ss = 0.0f;
            float ink_cost1 = 0.0f;
            float ink_cost2 = 0.0f;
            float knife_cost = 0.0f;
            float process1_cost_cms = 0.0f;
            float process2_cost_cms = 0.0f;

            float process1_cost_ss = 0.0f;
            float process2_cost_ss = 0.0f;

            float inspection_cost_cms = 0.0f;
            float inspection_cost_ss = 0.0f;
            float loss1_cms = 0.0f;
            float loss2_cms = 0.0f;
            float loss1_ss = 0.0f;
            float loss2_ss = 0.0f;

            float packing_fee_cms = 0.0f;
            float management_fee_cms = 0.0f;
            float delivery_fee_cms = 0.0f;
            float profit_value_cms = 0.0f;

            float packing_fee_ss = 0.0f;
            float management_fee_ss = 0.0f;
            float delivery_fee_ss = 0.0f;
            float profit_value_ss = 0.0f;


            float mc_rate_cms = 0.0f;
            float mc_rate_ss = 0.0f;


            float product_price_cms = 0.0f;
            float product_price_ss = 0.0f;

            DataGridViewRow row1 = dataGridView1.CurrentRow;            

            float pd = float.Parse(row1.Cells["PD"].Value.ToString());
            float cavity = float.Parse(row1.Cells["G_C"].Value.ToString())*float.Parse(row1.Cells["G_C_R"].Value.ToString());
            int step = int.Parse(row1.Cells["PROD_DIECUT_STEP"].Value.ToString());
            int print_times = int.Parse(row1.Cells["PROD_PRINT_TIMES"].Value.ToString());
            int knife_lifecycle = int.Parse(row1.Cells["KNIFE_LIFECYCLE"].Value.ToString());
            // knife_cost = float.Parse(row1.Cells["KNIFE_PRICE"].Value.ToString())*step/((float)knife_lifecycle*cavity);
            knife_cost = float.Parse(textBox26.Text) * step / ((float)knife_lifecycle * cavity)/22650;
            textBox3.Text = knife_cost.ToString();
            textBox14.Text = knife_cost.ToString();
                      


            if (dataGridView2.Rows.Count > 1)
            {
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    float single_mat_cost_cms = 0.0f;
                    float single_mat_cost_ss = 0.0f;

                    if (!row.IsNewRow)
                    {
                        float roll_price_cms = 0.0f;
                        float roll_price_ss = 0.0f;

                        float m_qty = float.Parse(row.Cells["M_QTY"].Value.ToString());
                        float roll_length = float.Parse(row.Cells["MAT_ROLL_LENGTH"].Value.ToString());
                        int roll_master_width = int.Parse(row.Cells["MAT_MASTER_WIDTH"].Value.ToString());
                       
                        int roll_cutwidth = int.Parse(row.Cells["MAT_CUTWIDTH"].Value.ToString());
                        if (roll_master_width == -1) roll_master_width = roll_cutwidth;

                        string KDYN = row.Cells["IMPORT_CAT"].Value.ToString();

                        float m_cms_price = (KDYN=="KD" ? (KDrate+1)*float.Parse(row.Cells["M_CMS_PRICE"].Value.ToString()): float.Parse(row.Cells["M_CMS_PRICE"].Value.ToString()));
                        float m_ss_price = (KDYN == "KD" ? (KDrate + 1) * float.Parse(row.Cells["M_SS_PRICE"].Value.ToString()) : float.Parse(row.Cells["M_SS_PRICE"].Value.ToString())) ;
                        float m_slitting_price = float.Parse(row.Cells["M_SLITTING_PRICE"].Value.ToString());
                        m_cms_price += m_slitting_price;

                        float left_over = (float) ((float)roll_master_width / (float)roll_cutwidth * 1.0f) - (roll_master_width / roll_cutwidth);
                        int roll_qty = (int)(roll_master_width / roll_cutwidth);
                        float average_loss = roll_cutwidth * left_over / (1.0f*roll_qty);

                        roll_price_cms = m_cms_price * (roll_cutwidth + average_loss) * roll_length / 1000.0f;
                        roll_price_ss = m_ss_price * (roll_cutwidth + average_loss) * roll_length / 1000.0f;

                        single_mat_cost_cms = m_qty*(float)(roll_price_cms / (roll_length * 1000 / pd*cavity));
                        single_mat_cost_ss = m_qty*(float)(roll_price_ss / (roll_length * 1000 / pd*cavity));

                        row.Cells["REMARK"].Value = (roll_cutwidth + average_loss).ToString();
                    }
                    material_cost_cms += single_mat_cost_cms;
                    material_cost_ss += single_mat_cost_ss;
                }
                textBox2.Text = material_cost_cms.ToString();
                textBox13.Text = material_cost_ss.ToString();
            }


            string product_size = (pd < 100 ? "SMALL" : pd < 200 ? "MEDIUM" : "LARGE");
            string EQ1 = row1.Cells["EQ1"].Value.ToString();
            string EQ2 = row1.Cells["EQ2"].Value.ToString();
            string PROD_TYPE = row1.Cells["PROD_TYPE"].Value.ToString();
            DataTable config1 = pro.getConfig(PROD_TYPE, product_size, EQ1, (EQ1 == "FR" || EQ1 == "SR" ? 1 : step));
            DataTable config2 = pro.getConfig(PROD_TYPE, product_size, EQ2, (EQ2 == "FR" || EQ2 == "SR" ? 1 : step));


            if (config1.Rows.Count > 0)
            {
                float phinhancong_khauhao = hesonhancong * float.Parse(config1.Rows[0]["LABOR_DEPRE_COST"].ToString());
                float phikiemtra = float.Parse(config1.Rows[0]["INSPECTION_COST"].ToString());
                float phimuc = print_times*float.Parse(config1.Rows[0]["INK_COST"].ToString());
                float tileNG = float.Parse(config1.Rows[0]["NG_RATE"].ToString());
                ink_cost1 = (EQ1 == "FR" || EQ1 == "SR" || EQ2 == "FR" || EQ2 == "SR" ? phimuc : 0);
              
                inspection_cost_cms = (ink_cost1 + knife_cost + material_cost_cms) * inspection_rate  + phikiemtra;
                inspection_cost_ss = (ink_cost1 + knife_cost + material_cost_ss) * inspection_rate + phikiemtra;

                loss1_cms = tileNG * ((ink_cost1 + knife_cost + material_cost_cms) + (phinhancong_khauhao/cavity));
                loss1_ss = tileNG * ((ink_cost1 + knife_cost + material_cost_ss)+(phinhancong_khauhao/cavity));

                process1_cost_cms = loss1_cms + (phinhancong_khauhao / cavity);
                process1_cost_ss = loss1_ss + (phinhancong_khauhao / cavity);


              //  MessageBox.Show("Inkcost = " + ink_cost1);
              //  MessageBox.Show($"phi nhan cong : {phinhancong_khauhao}, phi kiem tra : {phikiemtra}, phimuc: {phimuc}, tileNG: {tileNG}");
              //  MessageBox.Show("Loss1_cms = " + loss1_cms);
             //   MessageBox.Show("Loss1_samsung = " + loss1_ss);
              //  MessageBox.Show("Phi nhan cong /1ea: " + phinhancong_khauhao / cavity);

            }

            if (config2.Rows.Count > 0)
            {
                float phinhancong_khauhao = hesonhancong * float.Parse(config2.Rows[0]["LABOR_DEPRE_COST"].ToString());
                float phikiemtra = float.Parse(config2.Rows[0]["INSPECTION_COST"].ToString());
                float phimuc = float.Parse(config2.Rows[0]["INK_COST"].ToString());
                float tileNG = float.Parse(config2.Rows[0]["NG_RATE"].ToString());

                loss2_cms = tileNG *(process1_cost_cms + material_cost_cms +knife_cost +ink_cost1 + (phinhancong_khauhao / cavity));
                loss2_ss = tileNG *( process1_cost_ss + material_cost_ss + knife_cost+ ink_cost1 + (phinhancong_khauhao / cavity));

                process2_cost_cms =loss2_cms + (phinhancong_khauhao / cavity);
                process2_cost_ss = loss2_ss + (phinhancong_khauhao / cavity);

               // MessageBox.Show("Ti le Ng2= " + tileNG);
              //  MessageBox.Show("Loss2_cms = " + loss2_cms);
              //  MessageBox.Show("Loss2_samsung = " + loss2_ss);
              //  MessageBox.Show("Phi nhan cong /1ea: " + phinhancong_khauhao / cavity);


            }
            float sub_total_cms = knife_cost + ink_cost1+ material_cost_cms + process1_cost_cms + process2_cost_cms + inspection_cost_cms;
            float sub_total_ss = knife_cost + ink_cost1 +  material_cost_ss + process1_cost_ss + process2_cost_ss + inspection_cost_ss;



            management_fee_cms = sub_total_cms * management_rate;
            management_fee_ss = sub_total_ss * management_rate;

            packing_fee_cms = packing_rate * sub_total_cms;
            packing_fee_ss = packing_rate* sub_total_ss;

            sub_total_cms += management_fee_cms;
            sub_total_ss += management_fee_ss;


            delivery_fee_cms = sub_total_cms * delivery_rate;
            delivery_fee_ss = sub_total_ss * delivery_rate;
            sub_total_cms += delivery_fee_cms;
            sub_total_ss += delivery_fee_ss;

            profit_value_cms = sub_total_cms * profit_rate;
            profit_value_ss = sub_total_ss * profit_rate;


            product_price_cms = (ink_cost1+ knife_cost + material_cost_cms) + process1_cost_cms + process2_cost_cms +  management_fee_cms + packing_fee_cms + inspection_cost_cms + delivery_fee_cms + profit_value_cms;

            product_price_ss = ink_cost1 +  knife_cost + material_cost_ss + process1_cost_ss + process2_cost_ss +  management_fee_ss + packing_fee_ss + inspection_cost_ss +  delivery_fee_ss + profit_value_ss;


            mc_rate_cms = (ink_cost1 + knife_cost + material_cost_cms + loss1_cms + loss2_cms) / product_price_cms * 100;
            mc_rate_ss = (ink_cost1 + knife_cost + material_cost_ss + loss1_ss + loss2_ss) / product_price_ss * 100;

            textBox4.Text = process1_cost_cms.ToString();
            textBox5.Text = process2_cost_cms.ToString();
            textBox6.Text = inspection_cost_cms.ToString();
            textBox8.Text = packing_fee_cms.ToString();
            textBox9.Text = delivery_fee_cms.ToString();
            textBox28.Text = management_fee_cms.ToString();
            textBox10.Text = profit_value_cms.ToString();
            textBox11.Text = mc_rate_cms.ToString();
            textBox12.Text = product_price_cms.ToString();
            textBox7.Text = ink_cost1.ToString();

            textBox15.Text = process1_cost_ss.ToString();
            textBox16.Text = process2_cost_ss.ToString();
            textBox17.Text = inspection_cost_ss.ToString();
            textBox19.Text = packing_fee_ss.ToString();
            textBox20.Text = delivery_fee_ss.ToString();
            textBox29.Text = management_fee_ss.ToString();
            textBox21.Text = profit_value_ss.ToString();
            textBox22.Text = mc_rate_ss.ToString();
            textBox23.Text = product_price_ss.ToString();
            textBox18.Text = ink_cost1.ToString();

            try
            {
                // BEP calculation
                float BEP_mat_cost = material_cost_cms;
                float BEP_ink_cost = BEP_mat_cost * (print_times == 0 ? 0 : 1) * 0.012f;
                float BEP_sub_mat_cost = BEP_mat_cost * 0.04f;

                float BEP_labor_cost_unit = 35000.0f / 1150.0f / 8.0f; //3.8 USD
                float BEP_sx_man_power = float.Parse(row1.Cells["PROD_MANPOWER"].Value.ToString());
                float BE_kt_man_power = float.Parse(row1.Cells["INSPECT_MANPOWER"].Value.ToString());
                float BEP_sx_NG_RATE = float.Parse(row1.Cells["BEP_PROD_NG_RATE"].Value.ToString());
                float BEP_kt_NG_RATE = float.Parse(row1.Cells["BEP_INSP_NG_RATE"].Value.ToString());
                float BEP_1hour_prod_qty = float.Parse(row1.Cells["BEP_1HOUR_PROD_QTY"].Value.ToString());
                float depre_cost1 = 0.0f;
                float depre_cost2 = 0.0f;


                DataTable bepConfig1 = pro.getBEPConfig(EQ1);
                DataTable bepConfig2 = pro.getBEPConfig(EQ2);
                if (bepConfig1.Rows.Count > 0)
                {
                    depre_cost1 = (EQ1 == "FR" || EQ1 == "SR" ? print_times : step) * float.Parse(bepConfig1.Rows[0]["DEPRE_COST"].ToString());
                }
                if (bepConfig2.Rows.Count > 0)
                {
                    depre_cost2 = (EQ2 == "FR" || EQ2 == "SR" ? print_times : step) * float.Parse(bepConfig2.Rows[0]["DEPRE_COST"].ToString());
                }

                float BEP_main_sub_mat_cost = BEP_mat_cost + BEP_ink_cost + BEP_sub_mat_cost; // (1) + (2)
                float BEP_labor_cost = BEP_labor_cost_unit * (BEP_sx_man_power + BE_kt_man_power) / BEP_1hour_prod_qty; //(3)
                float BEP_Depre_cost = (depre_cost1 + depre_cost2) / BEP_1hour_prod_qty; //(4)
                float BEP_lost_cost = (BEP_sx_NG_RATE + BEP_kt_NG_RATE) * (BEP_main_sub_mat_cost + BEP_Depre_cost + BEP_labor_cost); //(c)
                float BEP_chetaogoc = BEP_main_sub_mat_cost + BEP_labor_cost + BEP_Depre_cost + BEP_lost_cost; //(d)
                float BEP_chetaogiantiep = (BEP_chetaogoc - BEP_Depre_cost) * 0.078f; // (e)
                float BEP_price = BEP_chetaogoc + BEP_chetaogiantiep; // (g) = (d) + (e) 
                float BEP_profit = BEP_price * 0.07f; // (f)
                float BEP_Target_price = BEP_price + BEP_profit;

                textBox32.Text = BEP_main_sub_mat_cost.ToString(); // phi vat lieu
                textBox34.Text = (BEP_Depre_cost + BEP_labor_cost).ToString(); // phi gia cong

                textBox33.Text = BEP_lost_cost.ToString(); // chi phi loss
                textBox37.Text = BEP_chetaogoc.ToString(); // chi phi truc tiep
                textBox38.Text = BEP_chetaogiantiep.ToString(); // chi phi gian tiep

                textBox39.Text = BEP_price.ToString(); // gia BEP
                textBox40.Text = BEP_profit.ToString(); // loi nhuan
                textBox41.Text = BEP_Target_price.ToString(); // target 

                //MessageBox.Show($"KHAU HAO = {BEP_Depre_cost} , NHAN CONG =  {BEP_labor_cost}");

            }
            catch (FormatException ex)
            {
                MessageBox.Show("Code này chưa được nhập thông tin tính BEP, sẽ k tính BEP");
            }
            catch(Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.ToString());
            }


           
        }
        private void button2_Click(object sender, EventArgs e)
        {
            tinhgia(1);
            MessageBox.Show("Tinh giá thành công");
        }

        private void TinhBaoGia_Load(object sender, EventArgs e)
        {
            int h = Screen.PrimaryScreen.WorkingArea.Height;
            int w = Screen.PrimaryScreen.WorkingArea.Width;
            this.ClientSize = new Size(w, h);

            checkBox1.Checked = true;
            setconfig("TSP");
           

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

            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;


        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                float MATERIAL_COST_CMS, PROCESS_COST_CMS, OTHER_COST_CMS, PROFIT_VALUE_CMS, MCR_CMS, MATERIAL_COST_SS, PROCESS_COST_SS, OTHER_COST_SS, PROFIT_VALUE_SS, MCR_SS, PRODUCT_CMSPRICE, PRODUCT_SSPRICE;
                DataGridViewRow row = dataGridView1.CurrentRow;

                string G_CODE = row.Cells["G_CODE"].Value.ToString();

                MATERIAL_COST_CMS = float.Parse(textBox2.Text) + float.Parse(textBox3.Text) + float.Parse(textBox7.Text);
                PROCESS_COST_CMS = float.Parse(textBox4.Text) + float.Parse(textBox5.Text);
                OTHER_COST_CMS = float.Parse(textBox6.Text) + float.Parse(textBox8.Text) + float.Parse(textBox9.Text) + float.Parse(textBox28.Text);
                PROFIT_VALUE_CMS = float.Parse(textBox10.Text);
                MCR_CMS = float.Parse(textBox11.Text);

                MATERIAL_COST_SS = float.Parse(textBox13.Text) + float.Parse(textBox14.Text) + float.Parse(textBox18.Text);
                PROCESS_COST_SS = float.Parse(textBox15.Text) + float.Parse(textBox16.Text);
                OTHER_COST_SS = float.Parse(textBox17.Text) + float.Parse(textBox19.Text) + float.Parse(textBox20.Text) + float.Parse(textBox29.Text);
                PROFIT_VALUE_SS = float.Parse(textBox21.Text);
                MCR_SS = float.Parse(textBox22.Text);

                PRODUCT_CMSPRICE = float.Parse(textBox12.Text);
                PRODUCT_SSPRICE = float.Parse(textBox23.Text);

                string updatevalue = $" SET MATERIAL_COST_CMS ='{MATERIAL_COST_CMS}',PROCESS_COST_CMS='{PROCESS_COST_CMS}',OTHER_COST_CMS='{OTHER_COST_CMS}',PROFIT_VALUE_CMS='{PROFIT_VALUE_CMS}',MCR_CMS='{MCR_CMS}',MATERIAL_COST_SS='{MATERIAL_COST_SS}',PROCESS_COST_SS='{PROCESS_COST_SS}',OTHER_COST_SS='{OTHER_COST_SS}',PROFIT_VALUE_SS='{PROFIT_VALUE_SS}',MCR_SS='{MCR_SS}', PRODUCT_CMSPRICE='{PRODUCT_CMSPRICE}', PRODUCT_SSPRICE='{PRODUCT_SSPRICE}' WHERE G_CODE='{G_CODE}'";

                //MessageBox.Show(G_CODE);

                ProductBLL pro = new ProductBLL();

                try
                {
                    pro.updatebaogiaM100(updatevalue);
                    MessageBox.Show("Lưu giá lên hệ thống thành công !");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: Bấm tính giá rồi mới lưu được giá !");
            }
           



        }
        public void setconfig(string type)
        {
            switch(type)
            {
                case "TSP":
                    textBox24.Text = "10"; //profit rate
                    textBox25.Text = "1";    //delivery_rate     
                    textBox30.Text = "8"; //management_rate
                    textBox35.Text = "15"; //KD-rate
                    textBox36.Text = "0"; //packing_rate
                    textBox31.Text = "2"; //inspection_rate
                    textBox26.Text = "4000000"; //knife price
                    textBox27.Text = "1"; //he so nhan cong

                    break;

                case "OLED":
                    textBox24.Text = "10"; //profit rate
                    textBox25.Text = "1";    //delivery_rate     
                    textBox30.Text = "15"; //management_rate
                    textBox35.Text = "15"; //KD-rate
                    textBox36.Text = "0"; //packing_rate
                    textBox31.Text = "2"; //inspection_rate
                    textBox26.Text = "5000000"; //knife price
                    textBox27.Text = "1"; //he so nhan cong
                    break;

                case "LABEL":
                    textBox24.Text = "10"; //profit rate
                    textBox25.Text = "1";    //delivery_rate     
                    textBox30.Text = "8"; //management_rate
                    textBox35.Text = "15"; //KD-rate
                    textBox36.Text = "0"; //packing_rate
                    textBox31.Text = "2"; //inspection_rate
                    textBox26.Text = "4000000"; //knife price
                    textBox27.Text = "1"; //he so nhan cong
                    break;

                default:
                    textBox24.Text = "10"; //profit rate
                    textBox25.Text = "1";    //delivery_rate     
                    textBox30.Text = "8"; //management_rate
                    textBox35.Text = "15"; //KD-rate
                    textBox36.Text = "0"; //packing_rate
                    textBox31.Text = "2"; //inspection_rate
                    textBox26.Text = "4000000"; //knife price
                    textBox27.Text = "1"; //he so nhan cong
                    break;
            }
           
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if(checkBox1.Checked == true)
            {
                textBox24.Enabled = false;
                textBox25.Enabled = false;
                textBox30.Enabled = false;
                textBox35.Enabled = false;
                textBox36.Enabled = false;
                textBox31.Enabled = false;
                textBox26.Enabled = false;
                textBox27.Enabled = false;
            }
            else
            {
                textBox24.Enabled = true;
                textBox25.Enabled = true;
                textBox30.Enabled = true;
                textBox35.Enabled = true;
                textBox36.Enabled = true;
                textBox31.Enabled = true;
                textBox26.Enabled = true;
                textBox27.Enabled = true;
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
          
        }

        private void checkVàUpdateThôngTinLiệuToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MaterialInfo mtrif = new MaterialInfo();
            mtrif.Show();
        }

        private void checkVàUpdateThôngTinBOMToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CheckBOMGia ckb = new CheckBOMGia();
            ckb.Show();
        }

        private void bảngGiáToolStripMenuItem_Click(object sender, EventArgs e)
        {
            BangGia bg = new BangGia();
            bg.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (dataGridView2 == null || dataGridView2.Rows.Count == 0)
            {
                MessageBox.Show("Tra bom trước");
            }
            else
            {
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    if (!row.IsNewRow)
                    {
                        if (row.Cells["CUST_CD"].Value.ToString() == "") row.Cells["CUST_CD"].Style.BackColor = Color.Red;
                        if (row.Cells["IMPORT_CAT"].Value.ToString() == "") row.Cells["IMPORT_CAT"].Style.BackColor = Color.Red;
                        if (row.Cells["M_CMS_PRICE"].Value.ToString() == "0") row.Cells["M_CMS_PRICE"].Style.BackColor = Color.Red;
                        if (row.Cells["M_SS_PRICE"].Value.ToString() == "0") row.Cells["M_SS_PRICE"].Style.BackColor = Color.Red;
                        if (row.Cells["USAGE"].Value.ToString() == "") row.Cells["USAGE"].Style.BackColor = Color.Red;
                        if (row.Cells["MAT_MASTER_WIDTH"].Value.ToString() == "0") row.Cells["MAT_MASTER_WIDTH"].Style.BackColor = Color.Red;
                        if (row.Cells["MAT_CUTWIDTH"].Value.ToString() == "0") row.Cells["MAT_CUTWIDTH"].Style.BackColor = Color.Red;
                        if (row.Cells["MAT_ROLL_LENGTH"].Value.ToString() == "0") row.Cells["MAT_ROLL_LENGTH"].Style.BackColor = Color.Red;
                        if (row.Cells["MAT_ROLL_LENGTH"].Value.ToString() == "0") row.Cells["MAT_ROLL_LENGTH"].Style.BackColor = Color.Red;
                    }
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();

                var selectedRows = dataGridView2.SelectedRows
                .OfType<DataGridViewRow>()
                .Where(row => !row.IsNewRow)
                .ToArray();

                foreach (var row in selectedRows)
                {
                    string
                        BOM_ID = row.Cells["BOM_ID"].Value.ToString(),
                        M_NAME = row.Cells["M_NAME"].Value.ToString();
                    dt = pro.getMaterialInfo(M_NAME, "ok");
                    if (dt.Rows.Count > 0)
                    {
                        row.Cells["CUST_CD"].Value = dt.Rows[0]["CUST_CD"];
                        row.Cells["M_CMS_PRICE"].Value = dt.Rows[0]["CMSPRICE"];
                        row.Cells["M_SS_PRICE"].Value = dt.Rows[0]["SSPRICE"];
                        row.Cells["M_SLITTING_PRICE"].Value = dt.Rows[0]["SLITTING_PRICE"];
                        row.Cells["MAT_MASTER_WIDTH"].Value = dt.Rows[0]["MASTER_WIDTH"];
                        row.Cells["MAT_ROLL_LENGTH"].Value = dt.Rows[0]["ROLL_LENGTH"];

                        row.Cells["CUST_CD"].Style.BackColor = Color.LightGray;
                        row.Cells["M_CMS_PRICE"].Style.BackColor = Color.LightGray;
                        row.Cells["M_SS_PRICE"].Style.BackColor = Color.LightGray;
                        row.Cells["M_SLITTING_PRICE"].Style.BackColor = Color.LightGray;
                        row.Cells["MAT_MASTER_WIDTH"].Style.BackColor = Color.LightGray;
                        row.Cells["MAT_ROLL_LENGTH"].Style.BackColor = Color.LightGray;

                    }

                    //pro.updateMaterial(updateValue);
                }
                dataGridView1.ClearSelection();
                MessageBox.Show("Cập nhật bom giá hoàn thành !");

            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi update Material: " + ex.ToString());
            }
        }
        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                ProductBLL pro = new ProductBLL();
                DataTable dt = new DataTable();

                var selectedRows = dataGridView2.SelectedRows
                .OfType<DataGridViewRow>()
                .Where(row => !row.IsNewRow)
                .ToArray();

                foreach (var row in selectedRows)
                {
                    string
                        BOM_ID = row.Cells["BOM_ID"].Value.ToString(),
                        CUST_CD = row.Cells["CUST_CD"].Value.ToString(),
                        M_CMS_PRICE = row.Cells["M_CMS_PRICE"].Value.ToString(),
                        M_SS_PRICE = row.Cells["M_SS_PRICE"].Value.ToString(),
                        M_SLITTING_PRICE = row.Cells["M_SLITTING_PRICE"].Value.ToString(),
                        MAT_MASTER_WIDTH = row.Cells["MAT_MASTER_WIDTH"].Value.ToString(),
                        MAT_ROLL_LENGTH = row.Cells["MAT_ROLL_LENGTH"].Value.ToString(),
                        USAGE = row.Cells["USAGE"].Value.ToString(),
                        MAT_CUTWIDTH = row.Cells["MAT_CUTWIDTH"].Value.ToString(),
                        M_QTY = row.Cells["M_QTY"].Value.ToString();

                    string updatevalue = $" SET CUST_CD='{CUST_CD}', M_CMS_PRICE='{M_CMS_PRICE}', M_SS_PRICE='{M_SS_PRICE}', M_SLITTING_PRICE='{M_SLITTING_PRICE}', MAT_MASTER_WIDTH='{MAT_MASTER_WIDTH}', MAT_ROLL_LENGTH='{MAT_ROLL_LENGTH}', USAGE='{USAGE}', MAT_CUTWIDTH='{MAT_CUTWIDTH}', M_QTY='{M_QTY}' WHERE BOM_ID={BOM_ID}";

                    pro.updateBOM2(updatevalue);

                }
                dataGridView1.ClearSelection();
                MessageBox.Show("Update Material info thành công !");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi update Material: " + ex.ToString());
            }
        }

        private void thôngSốTínhGiáToolStripMenuItem_Click(object sender, EventArgs e)
        {
            BaoGiaConfig bgconfig = new BaoGiaConfig();
            bgconfig.Show();
        }

        private void groupBox5_Enter(object sender, EventArgs e)
        {

        }

        private void label48_Click(object sender, EventArgs e)
        {

        }

        private void textBox39_TextChanged(object sender, EventArgs e)
        {

        }

        private void nhậpThôngTinTínhBEPToolStripMenuItem_Click(object sender, EventArgs e)
        {
            BEP_Info bepinfo = new BEP_Info();
            bepinfo.Show();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                float BEP_MAT_COST, BEP_PROC_COST, BEP_TOTAL_LOSS, BEP_PRICE, BEP_PROFIT_VALUE, BEP_TARGET_PRICE;
                DataGridViewRow row = dataGridView1.CurrentRow;

                string G_CODE = row.Cells["G_CODE"].Value.ToString();

                BEP_MAT_COST = float.Parse(textBox32.Text) + float.Parse(textBox3.Text) + float.Parse(textBox7.Text);
                BEP_PROC_COST = float.Parse(textBox34.Text) + float.Parse(textBox5.Text);
                BEP_TOTAL_LOSS = float.Parse(textBox33.Text) + float.Parse(textBox8.Text) + float.Parse(textBox9.Text) + float.Parse(textBox28.Text);
                BEP_PRICE = float.Parse(textBox39.Text);
                BEP_PROFIT_VALUE = float.Parse(textBox40.Text);

                BEP_TARGET_PRICE = float.Parse(textBox41.Text); 

                string updatevalue = $" SET BEP_MAT_COST ='{BEP_MAT_COST}',BEP_PROC_COST='{BEP_PROC_COST}',BEP_TOTAL_LOSS='{BEP_TOTAL_LOSS}',BEP_PRICE='{BEP_PRICE}',BEP_PROFIT_VALUE='{BEP_PROFIT_VALUE}',BEP_TARGET_PRICE='{BEP_TARGET_PRICE}' WHERE G_CODE='{G_CODE}'";

                //MessageBox.Show(G_CODE);

                ProductBLL pro = new ProductBLL();

                try
                {
                    pro.updatebaogiaM100(updatevalue);
                    MessageBox.Show("Lưu giá BEP lên hệ thống thành công !");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: Bấm tính giá rồi mới lưu được giá !");
            }

        }

        private void bảngGiáTheoCodeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            BangGia bg = new BangGia();
            bg.Show();
        }

        private void bảngGiáFullBOMToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
    }
}
