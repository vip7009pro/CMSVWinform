using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoClick
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new NewYCSX());
            Application.Run(new Form2());
            //Application.Run(new QuanLyKhachHang());


            //Application.Run(new BEP_Info());

            //Application.Run(new Chart1());
            //Application.Run(new HomePage());
            //Application.Run(new NewCodeBom());
            //Application.Run(new YCSX_Manager());
        }
    }
}
