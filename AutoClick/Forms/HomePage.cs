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
    public partial class HomePage : RibbonForm
    {
        public HomePage()
        {
            InitializeComponent();
        }

        private void ribbon1_Click(object sender, EventArgs e)
        {

        }

        private void ribbonButton2_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("ok");
            Form1 frm1 = new Form1();
            frm1.TopLevel = false;
            frm1.AutoScroll = true;
            panel1.Controls.Add(frm1);
            
            //frm1.MdiParent = this;
            frm1.Show();  
            
        }

        private void ribbonButton3_Click(object sender, EventArgs e)
        {
            MessageBox.Show("ok");
        }
    }
}
