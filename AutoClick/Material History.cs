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
    public partial class Material_History : Form
    {
        public Material_History()
        {
            InitializeComponent();
        }
        public string ycsx_no = "";
        public void tra_Material_History()
        {
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dt = pro.Material_History(ycsx_no, "chitiet");
            dataGridView1.DataSource = dt;

            dt = pro.Material_History(ycsx_no, "chitiethon");
            dataGridView2.DataSource = dt;

            dt = pro.Material_History(ycsx_no, "chitiethonnua");
            dataGridView3.DataSource = dt;

            dt = pro.Material_History(ycsx_no, "lieuinput");
            dataGridView4.DataSource = dt;

        }
        private void Material_History_Load(object sender, EventArgs e)
        {
            this.dataGridView1.DefaultCellStyle.ForeColor = Color.Blue;
            this.dataGridView1.DefaultCellStyle.BackColor = Color.Beige;

            this.dataGridView2.DefaultCellStyle.ForeColor = Color.Blue;
            this.dataGridView2.DefaultCellStyle.BackColor = Color.Beige;

            this.dataGridView3.DefaultCellStyle.ForeColor = Color.Blue;
            this.dataGridView3.DefaultCellStyle.BackColor = Color.Beige;

            this.dataGridView4.DefaultCellStyle.ForeColor = Color.Blue;
            this.dataGridView4.DefaultCellStyle.BackColor = Color.Beige;

        }
    }
}
