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
using ZXing;
using ZXing.Common;
using System.IO;
using System.Drawing.Imaging;
using System.Drawing.Printing;


namespace AutoClick
{
    public partial class DESIGN_AMAZONE : Form
    {
        public DESIGN_AMAZONE()
        {
            InitializeComponent();
        }

        public string Login_ID = "";
        private void DESIGN_AMAZONE_Load(object sender, EventArgs e)
        {
            if (!System.Windows.Forms.SystemInformation.TerminalServerSession)
            {
                Type dgvType = dataGridView1.GetType();
                PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                  BindingFlags.Instance | BindingFlags.NonPublic);
                pi.SetValue(dataGridView1, true, null);
            }

            foreach (FontFamily oneFontFamily in FontFamily.Families)
            {
                comboBox2.Items.Add(oneFontFamily.Name);
            }
            foreach (string printerName in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
            {
                comboBox3.Items.Add(printerName);
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(comboBox1.SelectedItem != null && comboBox2.SelectedItem != null && comboBox4.SelectedItem != null && textBox2.Text != "")
            {
                int rowIndex = dataGridView1.Rows.Add();
                DataGridViewRow row = dataGridView1.Rows[rowIndex];
                string PHANLOAI_DT = comboBox1.SelectedItem.ToString();
                string FONT_NAME = comboBox2.SelectedItem.ToString();
                string FONT_SIZE = numericUpDown6.Value.ToString();
                string FONT_STYLE = comboBox4.SelectedItem.ToString();
                string CAVITY_IN = textBox2.Text;
                //FONT_STYLE = FONT_STYLE == "Regular" ? "R" : FONT_STYLE == "Itatic" ? "I" : FONT_STYLE == "Underline" ? "U" : "B";
                //MessageBox.Show(FONT_STYLE);
                row.Cells["DOITUONG_NAME"].Value = "";
                row.Cells["CAVITY_PRINT"].Value = CAVITY_IN;
                row.Cells["PHANLOAI_DT"].Value = PHANLOAI_DT;
                row.Cells["CAVITY_PRINT"].Value = textBox2.Text;
                row.Cells["FONT_STYLE"].Value = FONT_STYLE;
                row.Cells["FONT_NAME"].Value = FONT_NAME;
                row.Cells["FONT_SIZE"].Value = FONT_SIZE;

                row.Cells["POS_X"].Value = 0;
                row.Cells["POS_Y"].Value = 0;
                row.Cells["SIZE_W"].Value = 0;
                row.Cells["SIZE_H"].Value = 0;
                row.Cells["ROTATE"].Value = 0;
                row.Cells["REMARK"].Value = "";

                dataGridView1.CurrentCell = row.Cells[0];
            }
            else
            {
                MessageBox.Show("Chọn đủ thông số trước đi đã rồi mới thêm đối tượng");
            }
            
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {            
          
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ProductBLL pro = new ProductBLL();
            DataTable dtBOM = new DataTable();
            DataTable dtMASS = new DataTable();
            
            if(label6.Text != "")
            {
                dtBOM = pro.getAmazoneBOM_GCODE(label6.Text);
                dtMASS = pro.getMassAmazone_GCODE(label6.Text);              

                if (dtMASS.Rows.Count > 0)
                {
                    MessageBox.Show("Design đã chạy mass không sửa được design, hãy tạo design mới");
                }
                else
                {
                    if (dtBOM.Rows.Count > 0)
                    {
                        MessageBox.Show("Đã có code sử dụng design này, không sửa được");
                    }
                    else
                    {
                        for (int r = 0; r < dataGridView3.Rows.Count; r++)
                        {
                            DataGridViewRow row = dataGridView3.Rows[r];
                            if (!row.IsNewRow)
                            {
                                string G_CODE_MAU = label6.Text;
                                string DOITUONG_NO = row.Cells["DOITUONG_NO"].Value.ToString();
                                string DOITUONG_NAME = row.Cells["DOITUONG_NAME"].Value.ToString();
                                string PHANLOAI_DT = row.Cells["PHANLOAI_DT"].Value.ToString();
                                string CAVITY_PRINT = row.Cells["CAVITY_PRINT"].Value.ToString();
                                string GIATRI = row.Cells["GIATRI"].Value.ToString();
                                string FONT_NAME = row.Cells["FONT_NAME"].Value.ToString();
                                string FONT_SIZE = row.Cells["FONT_SIZE"].Value.ToString();
                                string FONT_STYLE = row.Cells["FONT_STYLE"].Value.ToString();
                                string POS_X = row.Cells["POS_X"].Value.ToString();
                                string POS_Y = row.Cells["POS_Y"].Value.ToString();
                                string SIZE_W = row.Cells["SIZE_W"].Value.ToString();
                                string SIZE_H = row.Cells["SIZE_H"].Value.ToString();
                                string ROTATE = row.Cells["ROTATE"].Value.ToString();
                                string REMARK = row.Cells["REMARK"].Value.ToString();
                                string DOITUONG_STT = row.Cells["DOITUONG_STT"].Value.ToString();
                                string EMPL_NO = Login_ID;
                                string updateValue = $"SET DOITUONG_STT='{DOITUONG_STT}', DOITUONG_NAME='{DOITUONG_NAME}', PHANLOAI_DT='{PHANLOAI_DT}', CAVITY_PRINT='{CAVITY_PRINT}', GIATRI='{GIATRI}', FONT_NAME='{FONT_NAME}',FONT_SIZE='{FONT_SIZE}',FONT_STYLE='{FONT_STYLE}',POS_X='{POS_X}',POS_Y='{POS_Y}',SIZE_W='{SIZE_W}',SIZE_H='{SIZE_H}',ROTATE='{ROTATE}',REMARK='{REMARK}',UPD_EMPL='{EMPL_NO}',UPD_DATE=GETDATE() WHERE G_CODE_MAU = '{G_CODE_MAU}' AND DOITUONG_NO='{DOITUONG_NO}'";
                                try
                                {
                                    pro.UpdateDESIGNAmazone(updateValue);
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.ToString());
                                }
                            }
                        }
                        MessageBox.Show("Update Design thành công");

                    }
                }


            }
            else
            {
                MessageBox.Show("Chọn code phôi trước đã");
            }


        }

        public void print_label(DataGridView dtgv, float offsetX, float offsetY)
        {
            printDocument1.Dispose();       

            printDocument1.PrintPage += (sender, e) =>
            {
                e.Graphics.Clear(Color.White);
                //MessageBox.Show("Rows = " + dtgv.Rows.Count);

                for (int r = 0; r < dtgv.Rows.Count; r++)
                {
                    DataGridViewRow row = dtgv.Rows[r];
                    if (!row.IsNewRow)
                    {                        
                        string PHANLOAI_DT = row.Cells["PHANLOAI_DT"].Value.ToString();
                        string FONT_NAME = row.Cells["FONT_NAME"].Value.ToString();
                        //string FONT_NAME = "Arial";
                        float FONT_SIZE = float.Parse(row.Cells["FONT_SIZE"].Value.ToString());
                        string FONT_STYLE = row.Cells["FONT_STYLE"].Value.ToString();

                        //float FONT_SIZE = 6.0f;
                        FONT_STYLE = FONT_STYLE == "B" ? "Bold" : FONT_STYLE == "I" ? "Italic" : FONT_STYLE == "U" ? "Underline" : "Regular";
                        
                        float POS_X_mm = float.Parse(row.Cells["POS_X"].Value.ToString())+offsetX;
                        float POS_Y_mm = float.Parse(row.Cells["POS_Y"].Value.ToString())+offsetY;
                        float SIZE_W_mm = float.Parse(row.Cells["SIZE_W"].Value.ToString());
                        float SIZE_H_mm = float.Parse(row.Cells["SIZE_H"].Value.ToString());
                        float ROTATE_DEGREE = float.Parse(row.Cells["ROTATE"].Value.ToString());

                        /* Conver mm to Point*/
                        float DPI = 600.0f;
                        float factor = 1.0f;

                        
                        int POS_X = (int) (POS_X_mm * factor);
                        int POS_Y = (int)(POS_Y_mm * factor);
                        int SIZE_W = (int)(SIZE_W_mm * factor);
                        int SIZE_H = (int)(SIZE_H_mm * factor);                       

                        string GIATRI = "";
                       

                        if (row.Cells["GIATRI"].Value == null)
                        {
                             GIATRI = "";
                        }
                        else
                        {
                             GIATRI = row.Cells["GIATRI"].Value.ToString();                        
                        }


                        //e.Graphics.RotateTransform(ROTATE_DEGREE);

                        e.Graphics.PageUnit = GraphicsUnit.Millimeter;

                        if (PHANLOAI_DT == "TEXT")
                        {
                            FontStyle fst = (FontStyle)Enum.Parse(typeof(FontStyle), FONT_STYLE, true);
                            Font font_dt = new Font(FONT_NAME, FONT_SIZE, fst);
                            e.Graphics.RotateTransform(ROTATE_DEGREE);
                            e.Graphics.DrawString(GIATRI, font_dt, Brushes.Black, POS_X_mm,POS_Y_mm);
                            e.Graphics.RotateTransform(-ROTATE_DEGREE);

                        }
                        else if (PHANLOAI_DT == "IMAGE")
                        {
                            string Dir = System.IO.Directory.GetCurrentDirectory();
                            string drawpath = Dir + "\\Logo\\" + GIATRI;
                            System.Drawing.Image imagex = System.Drawing.Image.FromFile(drawpath);
                            Bitmap bitMap = new Bitmap(imagex);
                            e.Graphics.RotateTransform(ROTATE_DEGREE);
                            e.Graphics.DrawImage(bitMap, POS_X_mm, POS_Y_mm, SIZE_W_mm, SIZE_H_mm);
                            e.Graphics.RotateTransform(-ROTATE_DEGREE);
                        }
                        else if (PHANLOAI_DT == "1D BARCODE")
                        {
                            /* Cau hinh Barcode 128*/
                            var barcodeWriter01 = new BarcodeWriter
                            {
                                Format = BarcodeFormat.CODE_128,
                                Options = new EncodingOptions
                                {
                                    Height = 20,
                                    Width = 40,
                                    Margin = 0,
                                    PureBarcode = true
                                },
                            };

                            using (var bitmap = barcodeWriter01.Write(GIATRI))
                            {
                                using (var stream = new MemoryStream())
                                {
                                    bitmap.Save(stream, ImageFormat.Png);
                                    var image = Image.FromStream(stream);
                                    float x = POS_X_mm;
                                    float y = POS_Y_mm;
                                    float width = SIZE_W_mm;
                                    float height = SIZE_H_mm;
                                    e.Graphics.RotateTransform(ROTATE_DEGREE);
                                    e.Graphics.DrawImage(image, x, y, width, height);
                                    e.Graphics.RotateTransform(-ROTATE_DEGREE);
                                }
                            }
                        }
                        else if (PHANLOAI_DT == "2D MATRIX")
                        {
                            /* Cau hinh Matrix*/
                            var barcodeWriter00 = new BarcodeWriter
                            {
                                Format = BarcodeFormat.DATA_MATRIX,
                                Options = new EncodingOptions
                                {
                                    Height = 300,
                                    Width = 300,
                                    Margin = 0
                                }
                            };

                            /* Thuc hien in matrix */

                            using (var bitmap = barcodeWriter00.Write(GIATRI))
                            {
                                using (var stream = new MemoryStream())
                                {
                                    bitmap.Save(stream, ImageFormat.Png);
                                    var image = Image.FromStream(stream);                                    
                                    float x = POS_X_mm;
                                    float y = POS_Y_mm;
                                    float width = SIZE_W_mm;
                                    float height = SIZE_H_mm;
                                    if(checkBox1.Checked == false)
                                    {
                                        e.Graphics.RotateTransform(ROTATE_DEGREE);
                                        e.Graphics.DrawImage(image, x, y, width, height);
                                        e.Graphics.RotateTransform(-ROTATE_DEGREE);
                                    }
                                    
                                }
                            }    
                        }
                        else if (PHANLOAI_DT == "2D QR CODE")
                        {
                            /* Cau hinh Matrix*/
                            var barcodeWriter00 = new BarcodeWriter
                            {
                                Format = BarcodeFormat.QR_CODE,
                                Options = new EncodingOptions
                                {
                                    Height = 300,
                                    Width = 300,
                                    Margin = 0
                                }
                            };

                            /* Thuc hien in matrix */

                            using (var bitmap = barcodeWriter00.Write(GIATRI))
                            {
                                using (var stream = new MemoryStream())
                                {
                                    bitmap.Save(stream, ImageFormat.Png);
                                    var image = Image.FromStream(stream);
                                    float x = POS_X_mm;
                                    float y = POS_Y_mm;
                                    float width = SIZE_W_mm;
                                    float height = SIZE_H_mm;
                                    if(checkBox1.Checked == false)
                                    {
                                        e.Graphics.RotateTransform(ROTATE_DEGREE);
                                        e.Graphics.DrawImage(image, x, y, width, height);
                                        e.Graphics.RotateTransform(-ROTATE_DEGREE);
                                    }
                                   
                                }
                            }
                        }
                        else if (PHANLOAI_DT == "CONTAINER")
                        {
                            Pen blackPen = new Pen(Color.Black, 0.01f);
                            e.Graphics.RotateTransform(ROTATE_DEGREE);
                            e.Graphics.DrawRectangle(blackPen, POS_X_mm, POS_Y_mm, SIZE_W_mm, SIZE_H_mm);
                            e.Graphics.RotateTransform(-ROTATE_DEGREE);
                        }

                    }
                }
            };
            //printDocument1.Print();
        }               
        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e, string matrix1, string matrix2)
        {           

        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try
                {  
                    /*
                    printPreviewControl1.Document = printDocument1;
                    printPreviewControl1.Zoom = 2.5;
                    print_label(dataGridView1);
                    */
                    
                }
                catch(Exception ex)
                {
                    MessageBox.Show("Lỗi: " + ex.ToString());
                }
                
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            /*
            if (dataGridView1.Columns.Contains("DOITUONG_NAME") == true) dataGridView1.Columns.Remove("DOITUONG_NAME");
            if (dataGridView1.Columns.Contains("PHANLOAI_DT") == true) dataGridView1.Columns.Remove("PHANLOAI_DT");
            if (dataGridView1.Columns.Contains("FONT_NAME") == true) dataGridView1.Columns.Remove("FONT_NAME");
            if (dataGridView1.Columns.Contains("FONT_SIZE") == true) dataGridView1.Columns.Remove("FONT_SIZE");
            if (dataGridView1.Columns.Contains("FONT_STYLE") == true) dataGridView1.Columns.Remove("FONT_STYLE");
            if (dataGridView1.Columns.Contains("POS_X") == true) dataGridView1.Columns.Remove("POS_X");
            if (dataGridView1.Columns.Contains("POS_Y") == true) dataGridView1.Columns.Remove("POS_Y");
            if (dataGridView1.Columns.Contains("SIZE_W") == true) dataGridView1.Columns.Remove("SIZE_W");
            if (dataGridView1.Columns.Contains("SIZE_H") == true) dataGridView1.Columns.Remove("SIZE_H");
            if (dataGridView1.Columns.Contains("REMARK") == true) dataGridView1.Columns.Remove("REMARK");
            if (dataGridView1.Columns.Contains("CAVITY_PRINT") == true) dataGridView1.Columns.Remove("CAVITY_PRINT");
            if (dataGridView1.Columns.Contains("ROTATE") == true) dataGridView1.Columns.Remove("ROTATE");
            if (dataGridView1.Columns.Contains("GIATRI") == true) dataGridView1.Columns.Remove("GIATRI");
            */


            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dt =pro.checkDESIGNAmazone("");
            dataGridView3.DataSource = dt;
            
        }

        private void button3_Click(object sender, EventArgs e)
        {  
            
                print_label(dataGridView3, float.Parse(textBox3.Text), float.Parse(textBox4.Text));
                printPreviewControl1.Document = printDocument1;
                printPreviewControl1.Zoom = 2.0;           
           
            
        }

        public void print_action()
        {
            if (comboBox3.SelectedItem != null)
            {
                string printername = comboBox3.SelectedItem.ToString();
                printDocument1.PrinterSettings.PrinterName = (printername);                
                printDocument1.DefaultPageSettings.PaperSize = new PaperSize("Custom", 300, 200);
                printDocument1.PrinterSettings.DefaultPageSettings.Margins.Top = 0;
                printDocument1.PrinterSettings.DefaultPageSettings.Margins.Bottom = 0;
                printDocument1.PrinterSettings.DefaultPageSettings.Margins.Right = 0;
                printDocument1.PrinterSettings.DefaultPageSettings.Margins.Left = 0;
                printDocument1.PrinterSettings.Copies = 1;
                printDocument1.PrinterSettings.DefaultPageSettings.Landscape = false;                
                printDocument1.Print();
            }
            else
            {
                MessageBox.Show("Chọn máy in trước");
            }
            
        }
        private void button5_Click(object sender, EventArgs e)
        {        
            for(int i=0;i<int.Parse(textBox5.Text);i++)
            {
                print_action();
            }
            
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {   
            /*
            print_label(dataGridView1);
            printPreviewControl1.Document = printDocument1;
            printPreviewControl1.Zoom = 2.0;
            */
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                ProductBLL pro = new ProductBLL();
                DataTable dt = pro.getcodephoiamazone(textBox1.Text);
                dataGridView2.DataSource = dt;                
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            dataGridView1.Columns.Clear();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            try{

                print_label(dataGridView1, float.Parse(textBox3.Text), float.Parse(textBox4.Text));
                printPreviewControl1.Document = printDocument1;
                printPreviewControl1.Zoom = 2.0;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
           
        }

        private void button8_Click(object sender, EventArgs e)
        {
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            if(label6.Text != "")
            {
                dt = pro.checkDESIGNAmazone(label6.Text);
                if(dt.Rows.Count >0)
                {
                    MessageBox.Show("Phôi này đã có design, ko thể thêm nữa, hãy sửa design của phôi này");
                }
                else
                {
                    for (int r = 0; r < dataGridView1.Rows.Count; r++)
                    {
                        DataGridViewRow row = dataGridView1.Rows[r];
                        if (!row.IsNewRow)
                        {
                            string G_CODE_MAU = label6.Text;
                            string DOITUONG_NO = r.ToString();
                            string DOITUONG_NAME = row.Cells["DOITUONG_NAME"].Value.ToString();
                            string PHANLOAI_DT = row.Cells["PHANLOAI_DT"].Value.ToString();
                            string CAVITY_PRINT = row.Cells["CAVITY_PRINT"].Value.ToString();
                            string GIATRI = row.Cells["GIATRI"].Value.ToString();
                            string FONT_NAME = row.Cells["FONT_NAME"].Value.ToString();
                            string FONT_SIZE = row.Cells["FONT_SIZE"].Value.ToString();
                            string FONT_STYLE = row.Cells["FONT_STYLE"].Value.ToString();
                            string POS_X = row.Cells["POS_X"].Value.ToString();
                            string POS_Y = row.Cells["POS_Y"].Value.ToString();
                            string SIZE_W = row.Cells["SIZE_W"].Value.ToString();
                            string SIZE_H = row.Cells["SIZE_H"].Value.ToString();
                            string ROTATE = row.Cells["ROTATE"].Value.ToString();
                            string REMARK = row.Cells["REMARK"].Value.ToString();
                            string EMPL_NO = Login_ID;

                            string insertValue = $"('002','{G_CODE_MAU}','{DOITUONG_NO}','{DOITUONG_NAME}' ,'{PHANLOAI_DT}','{CAVITY_PRINT}','{FONT_NAME}','{POS_X}','{POS_Y}','{SIZE_W}','{SIZE_H}','{ROTATE}','{REMARK}',GETDATE(), '{EMPL_NO}', GETDATE(), '{EMPL_NO}','{FONT_SIZE}','{FONT_STYLE}','{GIATRI}')";
                            try
                            {
                                pro.InsertDESIGNAmazone(insertValue);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.ToString());
                            }
                        }
                    }

                }



            }
            else
            {
                MessageBox.Show("Chọn code phôi trước");
            }
            
        }

        private void dataGridView2_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView3.Columns.Clear();
            DataGridViewRow row = dataGridView2.Rows[e.RowIndex];
            string G_CODE = row.Cells["G_CODE"].Value.ToString();
            label6.Text = G_CODE;
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dt = pro.checkDESIGNAmazone(G_CODE);
            dataGridView3.DataSource = dt;
        }

        private void dataGridView3_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            print_label(dataGridView3, float.Parse(textBox3.Text), float.Parse(textBox4.Text));
            printPreviewControl1.Document = printDocument1;
            printPreviewControl1.Zoom = 2.0;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            ProductBLL pro = new ProductBLL();
            DataTable dtBOM = new DataTable();
            DataTable dtMASS = new DataTable();
            if (label6.Text != "")
            {
                dtBOM = pro.getAmazoneBOM_GCODE(label6.Text);
                dtMASS = pro.getMassAmazone_GCODE(label6.Text);

                if (dtMASS.Rows.Count > 0)
                {
                    MessageBox.Show("Design đã chạy mass không xoá design được");
                }
                else
                {
                    if (dtBOM.Rows.Count > 0)
                    {
                        MessageBox.Show("Design đã được dùng vào BOM của code nào đó, không xoá design được");
                    }
                    else
                    {
                        pro.DeleteDESIGNAmazone(label6.Text);
                        MessageBox.Show("Xóa Design thành công");
                    }
                }
            }
            else
            {
                MessageBox.Show("Chọn code phôi trước đã");
            }
        }

        private void textBox3_Leave(object sender, EventArgs e)
        {
            print_label(dataGridView3, float.Parse(textBox3.Text), float.Parse(textBox4.Text));
            printPreviewControl1.Document = printDocument1;
            printPreviewControl1.Zoom = 2.0;
        }

        private void textBox4_Leave(object sender, EventArgs e)
        {
            print_label(dataGridView3, float.Parse(textBox3.Text), float.Parse(textBox4.Text));
            printPreviewControl1.Document = printDocument1;
            printPreviewControl1.Zoom = 2.0;
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            for(int i=0;i<15;i++)
            {
                print_label(dataGridView3, float.Parse(textBox3.Text), float.Parse(textBox4.Text));
                printPreviewControl1.Document = printDocument1;
                printPreviewControl1.Zoom = 2.0;
                print_action();

            }
           
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {

        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (!backgroundWorker1.IsBusy)
            {               
                backgroundWorker1.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Đang trong tiến trình khác, thử lại sau");
            }
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }
    }
}
