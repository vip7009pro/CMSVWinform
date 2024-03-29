﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using AutoUpdaterDotNET;

namespace AutoClick
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        public int currentver = 31;

        public void checkUpdate()
        {
            System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
            System.Diagnostics.FileVersionInfo fvi = System.Diagnostics.FileVersionInfo.GetVersionInfo(assembly.Location);
            AutoUpdater.CheckForUpdateEvent += AutoUpdaterOnCheckForUpdateEvent;
            string version = fvi.FileVersion;
            label3.Text = "Phiên bản: " + version;
            AutoUpdater.DownloadPath = "update";
            AutoUpdater.Start("http://14.160.33.94:3010/update/update.xml");

            System.Timers.Timer timer = new System.Timers.Timer
            {
                Interval = 1 * 60 * 1000,
                SynchronizingObject = this
            };
            timer.Elapsed += delegate
            {
                AutoUpdater.Start("http://14.160.33.94:3010/update/update.xml");
            };
            timer.Start();
        }

        private void AutoUpdaterOnCheckForUpdateEvent(UpdateInfoEventArgs args)
        {
            if (args.IsUpdateAvailable)
            {
                DialogResult dialogResult;
                dialogResult =
                        MessageBox.Show(
                            $@"Bạn ơi, phần mềm của bạn có phiên bản mới {args.CurrentVersion}. Phiên bản bạn đang sử dụng hiện tại  {args.InstalledVersion}. Bạn có muốn cập nhật phần mềm không?", @"Cập nhật phần mềm",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Information);

                if (dialogResult.Equals(DialogResult.Yes) || dialogResult.Equals(DialogResult.OK))
                {
                    try
                    {
                        if (AutoUpdater.DownloadUpdate(args))
                        {
                            Application.Exit();
                        }
                    }
                    catch (Exception exception)
                    {
                        MessageBox.Show(exception.Message, exception.GetType().ToString(), MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                    }
                }
            }
            else
            {
                /* MessageBox.Show(@"Phiên bản bạn đang sử dụng đã được cập nhật mới nhất.", @"Cập nhật phần mềm",
                     MessageBoxButtons.OK, MessageBoxIcon.Information); */
            }
        }



        public void login_fb()
        {

            string user = textBox1.Text;
            textBox2.PasswordChar = '*';
            string pass = textBox2.Text;
            if(user =="   ")
            {
                try
                {                    
                    string line;
                    // Read the file and display it line by line.  
                    System.IO.StreamReader file =
                        new System.IO.StreamReader("account.txt");
                    if ((line = file.ReadLine()) != null)
                    {   
                        //MessageBox.Show("Line content: " + line);
                        user = line;                        
                    }
                    if ((line = file.ReadLine()) != null)
                    {  
                        //MessageBox.Show("Line content: " + line);
                        pass = line;                        
                    }
                    file.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Có lỗi !\n" + ex.ToString());
                }
            }
            ProductBLL pr = new ProductBLL();
            DataTable dt = new DataTable();
            
            dt = pr.login(user, pass);
            if (dt.Rows.Count > 0)
            {
                //MessageBox.Show("Đăng nhập thành công");
                Form1 frm1 = new Form1();
                Form3 frm3 = new Form3();
                frm1.LoginID = user;
                frm3.loginIDfrm3 = user;
                frm1.Show();
                this.Hide();
            }
            else
            {
                MessageBox.Show("User hoặc password sai");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ProductBLL pro = new ProductBLL();           
            int lastver = pro.getVer();            
            if(currentver >= lastver)
            {
                login_fb();
            }
            else
            {
                MessageBox.Show("Last Ver = " + lastver + "\n Ver hien tai = " + currentver);
                MessageBox.Show("Phiên bản đã cũ,  Sẽ đưa bạn tới ver mới nhất, hãy tải về và ghi đè vào file cũ");
                System.Diagnostics.Process.Start("http://14.160.33.94/update/ERP2/lastest.zip");
            } 
        }

        private void Form2_KeyDown(object sender, KeyEventArgs e)
        {           
                if (e.KeyCode == Keys.Enter)
                {
                    login_fb();
                }            
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                ProductBLL pro = new ProductBLL();
                
                int lastver = pro.getVer();               

                if (currentver >= lastver)
                {
                    login_fb();
                }
                else
                {
                    MessageBox.Show("Last Ver = " + lastver + "\n Ver hien tai = " + currentver);
                    MessageBox.Show("Phiên bản đã cũ, hãy cập nhật, liên hệ ĐKA Hùng!");
                }
            }

        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                ProductBLL pro = new ProductBLL();
                
                int lastver = pro.getVer();                

                if (currentver >= lastver)
                {
                    login_fb();
                }
                else
                {
                    MessageBox.Show("Last Ver = " + lastver + "\n Ver hien tai = " + currentver);
                    MessageBox.Show("Phiên bản đã cũ, hãy cập nhật, liên hệ ĐKA Hùng!");
                }

            }
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            if (System.Diagnostics.Process.GetProcessesByName(System.IO.Path.GetFileNameWithoutExtension(System.Reflection.Assembly.GetEntryAssembly().Location)).Count() > 1) System.Diagnostics.Process.GetCurrentProcess().Kill();
            checkUpdate();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void gradientPanel1_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
