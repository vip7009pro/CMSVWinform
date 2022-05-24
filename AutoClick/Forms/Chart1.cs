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
            // TODO: This line of code loads data into the 'cMS_VINADataSet.ZTBPOTable' table. You can move, or remove it, as needed.
            this.zTBPOTableTableAdapter.Fill(this.cMS_VINADataSet.ZTBPOTable);

        }
    }
}
