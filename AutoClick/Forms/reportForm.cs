using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoClick
{
    public partial class reportForm : Form
    {
        public reportForm()
        {
            InitializeComponent();
        }

        public static DateTime FirstDayOfWeek(DateTime date)
        {
            DayOfWeek fdow = CultureInfo.CurrentCulture.DateTimeFormat.FirstDayOfWeek;
            int offset = fdow - date.DayOfWeek;
            DateTime fdowDate = date.AddDays(offset);
            return fdowDate;
        }

        public static DateTime LastDayOfWeek(DateTime date)
        {
            DateTime ldowDate = FirstDayOfWeek(date).AddDays(6);
            return ldowDate;
        }

        public static DateTime FridayofWeek(DateTime date)
        {
            DateTime ldowDate = FirstDayOfWeek(date).AddDays(5);
            return ldowDate;
        }


        public int GetLastWeekNumber(string dt)
        {
            DateTime dd = DateTime.Parse(dt);
            dd = dd.AddDays(1);
            CultureInfo ciCurr = CultureInfo.CurrentCulture;
            int weekNum = ciCurr.Calendar.GetWeekOfYear(dd, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
            return weekNum;
        }


        public void initdata()
        {
            DateTime tempdate = dateTimePicker1.Value;
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            DataTable dtPoBalance = new DataTable();
            DataTable dtPoBalanceType = new DataTable();
            DataTable dtPoBalanceTypeSS = new DataTable();

            DataTable dtPo = new DataTable();
            DataTable dtPoType = new DataTable();
            DataTable dtPoTypeSS = new DataTable();


            DataTable fcstthisweek = new DataTable();
            DataTable fcstlastweek = new DataTable();

            string startdate, enddate;
            DateTime FIX_ST_DATE = FirstDayOfWeek(dateTimePicker1.Value);
            DateTime FIX_EN_DATE = LastDayOfWeek(dateTimePicker1.Value);

            DateTime ST_DATE = FIX_ST_DATE;
            DateTime EN_DATE = FIX_EN_DATE;

            startdate = new Form1().STYMD2(ST_DATE.Year, ST_DATE.Month, ST_DATE.Day);
            enddate = new Form1().STYMD2(EN_DATE.Year, EN_DATE.Month, EN_DATE.Day);
            //MessageBox.Show("Start Date =  " + startdate);
            //MessageBox.Show("End Date =  " + enddate);

            //dt = pro.report_WeeklyPOByCustomer();
            for (int i = 0; i <= 10; i++)
            {
                
                dt = pro.report_WeeklyPOByCustomer2(startdate, enddate);               
                dtPo.Merge(dt);
                ST_DATE = ST_DATE.AddDays(-7);
                EN_DATE = EN_DATE.AddDays(-7);
                startdate = new Form1().STYMD2(ST_DATE.Year, ST_DATE.Month, ST_DATE.Day);
                enddate = new Form1().STYMD2(EN_DATE.Year, EN_DATE.Month, EN_DATE.Day);                
            }
                        

            dataGridView1.DataSource = dtPo;


            ST_DATE = FIX_ST_DATE;
            EN_DATE = FIX_EN_DATE;
            startdate = new Form1().STYMD2(ST_DATE.Year, ST_DATE.Month, ST_DATE.Day);
            enddate = new Form1().STYMD2(EN_DATE.Year, EN_DATE.Month, EN_DATE.Day);
            
            for (int i = 0; i <= 10; i++)
            {
                if(checkBox1.Checked ==true)
                {
                    dt = pro.report_WeeklyPOByTypeALL2SS(startdate, enddate);
                }
                else
                {
                    dt = pro.report_WeeklyPOByTypeALL2(startdate, enddate);
                }                              

                dtPoType.Merge(dt);
                ST_DATE = ST_DATE.AddDays(-7);
                EN_DATE = EN_DATE.AddDays(-7);
                startdate = new Form1().STYMD2(ST_DATE.Year, ST_DATE.Month, ST_DATE.Day);
                enddate = new Form1().STYMD2(EN_DATE.Year, EN_DATE.Month, EN_DATE.Day);
            }

            dataGridView5.DataSource = dtPoType;


           
            ST_DATE = FIX_ST_DATE;
            EN_DATE = FIX_EN_DATE;
            startdate = new Form1().STYMD2(ST_DATE.Year, ST_DATE.Month, ST_DATE.Day);
            enddate = new Form1().STYMD2(EN_DATE.Year, EN_DATE.Month, EN_DATE.Day);

            for (int i = 0; i <= 10; i++)
            {                
                dt = pro.report_WeeklyPOByTypeALL2SS(startdate, enddate);
                dtPoTypeSS.Merge(dt);
                ST_DATE = ST_DATE.AddDays(-7);
                EN_DATE = EN_DATE.AddDays(-7);
                startdate = new Form1().STYMD2(ST_DATE.Year, ST_DATE.Month, ST_DATE.Day);
                enddate = new Form1().STYMD2(EN_DATE.Year, EN_DATE.Month, EN_DATE.Day);
            }

            dataGridView7.DataSource = dtPoTypeSS;






            //dt = pro.report_WeeklyPOByType(new Form1().STYMD2(tempdate.Year, tempdate.Month, tempdate.Day));
            // dt = pro.report_WeeklyPOByType(new Form1().STYMD2(tempdate.Year, tempdate.Month, tempdate.Day));

            //dataGridView5.DataSource = dt;


            dt = pro.report_POBalanceByCustomer(new Form1().STYMD2(tempdate.Year, tempdate.Month, tempdate.Day));
            dtPoBalance.Merge(dt);
            tempdate = FridayofWeek(dateTimePicker1.Value.AddDays(-7));
            for (int i = 0; i <= 10; i++)
            {
                dt = pro.report_POBalanceByCustomer(new Form1().STYMD2(tempdate.Year, tempdate.Month, tempdate.Day));
                dtPoBalance.Merge(dt);
                tempdate = tempdate.AddDays(-7);
            }
            dataGridView2.DataSource = dtPoBalance;





            tempdate = dateTimePicker1.Value;
            dt = pro.report_POBalanceByType(new Form1().STYMD2(tempdate.Year, tempdate.Month, tempdate.Day));
            dtPoBalanceType.Merge(dt);
            tempdate = FridayofWeek(dateTimePicker1.Value.AddDays(-7));
            for (int i = 0; i <= 10; i++)
            {
                dt = pro.report_POBalanceByType(new Form1().STYMD2(tempdate.Year, tempdate.Month, tempdate.Day));                   
                dtPoBalanceType.Merge(dt);
                tempdate = tempdate.AddDays(-7);
            }

            dataGridView6.DataSource = dtPoBalanceType;

            tempdate = dateTimePicker1.Value;
            dt = pro.report_POBalanceByTypeSS(new Form1().STYMD2(tempdate.Year, tempdate.Month, tempdate.Day));
            dtPoBalanceTypeSS.Merge(dt);
            tempdate = FridayofWeek(dateTimePicker1.Value.AddDays(-7));
            for (int i = 0; i <= 10; i++)
            {    
                dt = pro.report_POBalanceByTypeSS(new Form1().STYMD2(tempdate.Year, tempdate.Month, tempdate.Day));
                dtPoBalanceTypeSS.Merge(dt);
                tempdate = tempdate.AddDays(-7);
            }
            dataGridView8.DataSource = dtPoBalanceTypeSS;




            int thisweeknum=0;
            int fcstyear=0;
            int fcstyear0=0;
            int lastfcstweeknum=0;
            int lastweeknum = 0;



            if(textBox1.Text == "" || textBox2.Text == "")
            {
                thisweeknum = new Form1().GetWeekNumber(dateTimePicker1.Value);
                fcstyear = dateTimePicker1.Value.Year;
                fcstyear0 = fcstyear;
                lastfcstweeknum = pro.getlastFCSTWeekNum(fcstyear);
               // MessageBox.Show("Last FCSt weeknum = " + lastfcstweeknum);

                if (lastfcstweeknum == -1) { fcstyear = fcstyear - 1; fcstyear0 = fcstyear; thisweeknum = 53; lastfcstweeknum = 52; }
                else if (lastfcstweeknum < thisweeknum) thisweeknum = lastfcstweeknum;

                if (thisweeknum == 1) { lastweeknum = 53; fcstyear0 = fcstyear - 1; } else { lastweeknum = thisweeknum - 1; }
            }
            else
            {
                fcstyear = int.Parse(textBox1.Text);
                thisweeknum = int.Parse(textBox2.Text);
                if (thisweeknum == 1) { fcstyear0 = fcstyear - 1;  lastweeknum = pro.getlastFCSTWeekNum(fcstyear0);  } else { lastweeknum = thisweeknum - 1; fcstyear0 = fcstyear; }
                //MessageBox.Show("fcst year =" + fcstyear);
                //MessageBox.Show("fcst year0 =" + fcstyear0);
                //MessageBox.Show("thisweeknum =" + thisweeknum);
                //MessageBox.Show("lastweeknum =" + lastweeknum);
            }

            //MessageBox.Show("This week and year: " + thisweeknum + "&" + fcstyear);
            //MessageBox.Show("Last week and year: " + lastweeknum + "&" + fcstyear0);

            fcstthisweek = pro.report_CustomerFcstByWeek(thisweeknum, fcstyear);
            fcstlastweek = pro.report_CustomerFcstByWeek(lastweeknum, fcstyear0);

            DataRow dr = fcstthisweek.NewRow();
            int colnum = fcstthisweek.Columns.Count;
            int rownum = fcstthisweek.Rows.Count;
            int sum = 0;
            for (int i = 2; i < colnum; i++)
            {
                for (int j = 0; j < rownum; j++)
                {
                    sum += int.Parse(fcstthisweek.Rows[j][fcstthisweek.Columns[i]].ToString());                    
                }
                dr[i] = sum;               
                sum = 0;               
            }
            dr[1] = "TOTAL";
            
            fcstthisweek.Rows.Add(dr);
           
            DataRow drlast = fcstlastweek.NewRow();
            int colnumlast = fcstlastweek.Columns.Count;
            int rownumlast = fcstlastweek.Rows.Count;
            sum = 0;
            for (int i = 2; i < colnumlast; i++)
            {
                for (int j = 0; j < rownumlast; j++)
                {
                    sum += int.Parse(fcstlastweek.Rows[j][fcstlastweek.Columns[i]].ToString());                    
                }
                drlast[i] = sum;
                sum = 0;
                //MessageBox.Show("SUM = " + sum);
            }
            drlast[1] = "TOTAL";
            fcstlastweek.Rows.Add(drlast);

            DataRow drdiff = fcstthisweek.NewRow();
            int colnumdiff = fcstthisweek.Columns.Count;
            int rownumdiff = fcstthisweek.Rows.Count;
            int diff = 0;
            for (int i = 2; i < colnum-1; i++)
            {
                diff = int.Parse(fcstthisweek.Rows[rownumdiff-1][fcstthisweek.Columns[i]].ToString()) - int.Parse(fcstlastweek.Rows[rownumdiff-1][fcstlastweek.Columns[i+1]].ToString());
                drdiff[i] = diff;                
            }
            drdiff[1] = "DIFFERENCE";
            fcstthisweek.Rows.Add(drdiff);


            DataRow drptram = fcstthisweek.NewRow();
            int colnumptram = fcstthisweek.Columns.Count;
            int rownumptram = fcstthisweek.Rows.Count;
            double ptram = 0.0;
            for (int i = 2; i < colnum - 1; i++)
            {
                int fcsttw = int.Parse(fcstthisweek.Rows[rownumdiff - 1][fcstthisweek.Columns[i]].ToString());
                int fcstlw = int.Parse(fcstlastweek.Rows[rownumdiff - 1][fcstlastweek.Columns[i + 1]].ToString());

                if(fcstlw !=0)
                {
                    ptram = (fcsttw- fcstlw) *1.0/ fcstlw *100.0 ;
                }
                else 
                {
                    if(fcsttw == 0)
                    {
                        ptram = 0;
                    }
                    else
                    {
                        ptram = 100;
                    }
                    
                }
                drptram[i] = ptram;
            }

            drptram[1] = "RATE";
            fcstthisweek.Rows.Add(drptram);

            DataRow drW = fcstthisweek.NewRow();
            drW[1] = "WEEK";
            drW[2] = thisweeknum+1;
            int weeknumcolumnum = fcstthisweek.Columns.Count;
            
            if (weeknumcolumnum >= 3)
            {
                for (int i = 3; i < weeknumcolumnum; i++)
                {   if(int.Parse(drW[i - 1].ToString())== GetLastWeekNumber(DateTime.Now.Year.ToString() + "-12-31"))
                    {
                        drW[i] = 1;
                    }
                    else
                    {
                        drW[i] = int.Parse(drW[i - 1].ToString()) + 1;
                    }                   

                }

            }

            fcstthisweek.Rows.InsertAt(drW, 0);



            DataRow drLW = fcstlastweek.NewRow();
            drLW[1] = "WEEK";
            //drLW[2] = (lastweeknum+1 > lastfcstweeknum) ? lastfcstweeknum : lastweeknum + 1;
            drLW[2] = (int.Parse(drW[2].ToString())-1)==0? lastfcstweeknum: (int.Parse(drW[2].ToString()) - 1);
            int weeknumcolumnumlast = fcstlastweek.Columns.Count;
            //DateTime tempdate = dateTimePicker1.Value;
            if (weeknumcolumnumlast >= 3)
            {
                for (int i = 3; i < weeknumcolumnumlast; i++)
                {
                    if (int.Parse(drLW[i - 1].ToString()) == lastfcstweeknum)
                    {
                        drLW[i] = 1;
                    }
                    else
                    {
                        drLW[i] = int.Parse(drLW[i - 1].ToString()) + 1;
                    }
                }
            }
            fcstlastweek.Rows.InsertAt(drLW, 0);
            fcstthisweek.Merge(fcstlastweek);

            /*
            for(int i=2;i<fcstthisweek.Columns.Count;i++)
            {
                fcstthisweek.Columns[i].DataType = typeof(double);
            }
            */

            dataGridView3.DataSource = fcstthisweek;
            setRowNumber(dataGridView1);
            setRowNumber(dataGridView2);
            setRowNumber(dataGridView6);
            setRowNumber(dataGridView5);
            setRowNumber(dataGridView3);

            formatWeeklyPOByCustomer(dataGridView1);
            formatWeeklyPOBalanceByCustomer(dataGridView2);
            
            formatWeeklyPOByType(dataGridView5);
            formatWeeklyPOBalanceByType(dataGridView6);
            formatWeeklyPOByType(dataGridView7);
            formatWeeklyPOBalanceByType(dataGridView8);
            if (dataGridView3.Rows.Count >=9)
            formatWeeklyFCSTByCustomer(dataGridView3);
            loadPOBalance();

        }

        public void initdata3()
        {
            DateTime tempdate = dateTimePicker1.Value;
            ProductBLL pro = new ProductBLL();             

            string startdate, enddate;
            DateTime FIX_ST_DATE = FirstDayOfWeek(dateTimePicker1.Value);
            DateTime FIX_EN_DATE = LastDayOfWeek(dateTimePicker1.Value);

            DateTime ST_DATE = FIX_ST_DATE;
            DateTime EN_DATE = FIX_EN_DATE;

            startdate = new Form1().STYMD2(ST_DATE.Year, ST_DATE.Month, ST_DATE.Day);
            enddate = new Form1().STYMD2(EN_DATE.Year, EN_DATE.Month, EN_DATE.Day);
            //MessageBox.Show("Start Date =  " + startdate);
            //MessageBox.Show("End Date =  " + enddate);

            //dt = pro.report_WeeklyPOByCustomer();
            for (int i = 0; i <= 10; i++)
            {

                dt = pro.report_WeeklyPOByCustomer2(startdate, enddate);
                dtPo.Merge(dt);
                ST_DATE = ST_DATE.AddDays(-7);
                EN_DATE = EN_DATE.AddDays(-7);
                startdate = new Form1().STYMD2(ST_DATE.Year, ST_DATE.Month, ST_DATE.Day);
                enddate = new Form1().STYMD2(EN_DATE.Year, EN_DATE.Month, EN_DATE.Day);
            }

            backgroundWorker1.ReportProgress(1);
            


            ST_DATE = FIX_ST_DATE;
            EN_DATE = FIX_EN_DATE;
            startdate = new Form1().STYMD2(ST_DATE.Year, ST_DATE.Month, ST_DATE.Day);
            enddate = new Form1().STYMD2(EN_DATE.Year, EN_DATE.Month, EN_DATE.Day);

            for (int i = 0; i <= 10; i++)
            {
                if (checkBox1.Checked == true)
                {
                    dt = pro.report_WeeklyPOByTypeALL2SS(startdate, enddate);
                }
                else
                {
                    dt = pro.report_WeeklyPOByTypeALL2(startdate, enddate);
                }

                dtPoType.Merge(dt);
                ST_DATE = ST_DATE.AddDays(-7);
                EN_DATE = EN_DATE.AddDays(-7);
                startdate = new Form1().STYMD2(ST_DATE.Year, ST_DATE.Month, ST_DATE.Day);
                enddate = new Form1().STYMD2(EN_DATE.Year, EN_DATE.Month, EN_DATE.Day);
            }
            backgroundWorker1.ReportProgress(2);
            



            ST_DATE = FIX_ST_DATE;
            EN_DATE = FIX_EN_DATE;
            startdate = new Form1().STYMD2(ST_DATE.Year, ST_DATE.Month, ST_DATE.Day);
            enddate = new Form1().STYMD2(EN_DATE.Year, EN_DATE.Month, EN_DATE.Day);

            for (int i = 0; i <= 10; i++)
            {
                dt = pro.report_WeeklyPOByTypeALL2SS(startdate, enddate);
                dtPoTypeSS.Merge(dt);
                ST_DATE = ST_DATE.AddDays(-7);
                EN_DATE = EN_DATE.AddDays(-7);
                startdate = new Form1().STYMD2(ST_DATE.Year, ST_DATE.Month, ST_DATE.Day);
                enddate = new Form1().STYMD2(EN_DATE.Year, EN_DATE.Month, EN_DATE.Day);
            }

            
            backgroundWorker1.ReportProgress(3);






            //dt = pro.report_WeeklyPOByType(new Form1().STYMD2(tempdate.Year, tempdate.Month, tempdate.Day));
            // dt = pro.report_WeeklyPOByType(new Form1().STYMD2(tempdate.Year, tempdate.Month, tempdate.Day));

            //dataGridView5.DataSource = dt;


            dt = pro.report_POBalanceByCustomer(new Form1().STYMD2(tempdate.Year, tempdate.Month, tempdate.Day));
            dtPoBalance.Merge(dt);
            tempdate = FridayofWeek(dateTimePicker1.Value.AddDays(-7));
            for (int i = 0; i <= 10; i++)
            {
                dt = pro.report_POBalanceByCustomer(new Form1().STYMD2(tempdate.Year, tempdate.Month, tempdate.Day));
                dtPoBalance.Merge(dt);
                tempdate = tempdate.AddDays(-7);
            }
            
            backgroundWorker1.ReportProgress(4);





            tempdate = dateTimePicker1.Value;
            dt = pro.report_POBalanceByType(new Form1().STYMD2(tempdate.Year, tempdate.Month, tempdate.Day));
            dtPoBalanceType.Merge(dt);
            tempdate = FridayofWeek(dateTimePicker1.Value.AddDays(-7));
            for (int i = 0; i <= 10; i++)
            {
                dt = pro.report_POBalanceByType(new Form1().STYMD2(tempdate.Year, tempdate.Month, tempdate.Day));
                dtPoBalanceType.Merge(dt);
                tempdate = tempdate.AddDays(-7);
            }

            
            backgroundWorker1.ReportProgress(5);

            tempdate = dateTimePicker1.Value;
            dt = pro.report_POBalanceByTypeSS(new Form1().STYMD2(tempdate.Year, tempdate.Month, tempdate.Day));
            dtPoBalanceTypeSS.Merge(dt);
            tempdate = FridayofWeek(dateTimePicker1.Value.AddDays(-7));
            for (int i = 0; i <= 10; i++)
            {
                dt = pro.report_POBalanceByTypeSS(new Form1().STYMD2(tempdate.Year, tempdate.Month, tempdate.Day));
                dtPoBalanceTypeSS.Merge(dt);
                tempdate = tempdate.AddDays(-7);
            }
            
            backgroundWorker1.ReportProgress(6);




            int thisweeknum = 0;
            int fcstyear = 0;
            int fcstyear0 = 0;
            int lastfcstweeknum = 0;
            int lastweeknum = 0;



            if (textBox1.Text == "" || textBox2.Text == "")
            {
                thisweeknum = new Form1().GetWeekNumber(dateTimePicker1.Value);
                fcstyear = dateTimePicker1.Value.Year;
                fcstyear0 = fcstyear;
                lastfcstweeknum = pro.getlastFCSTWeekNum(fcstyear);
                // MessageBox.Show("Last FCSt weeknum = " + lastfcstweeknum);

                if (lastfcstweeknum == -1) { fcstyear = fcstyear - 1; fcstyear0 = fcstyear; thisweeknum = 53; lastfcstweeknum = 52; }
                else if (lastfcstweeknum < thisweeknum) thisweeknum = lastfcstweeknum;

                if (thisweeknum == 1) { lastweeknum = 53; fcstyear0 = fcstyear - 1; } else { lastweeknum = thisweeknum - 1; }
            }
            else
            {
                fcstyear = int.Parse(textBox1.Text);
                thisweeknum = int.Parse(textBox2.Text);
                if (thisweeknum == 1) { fcstyear0 = fcstyear - 1; lastweeknum = pro.getlastFCSTWeekNum(fcstyear0); } else { lastweeknum = thisweeknum - 1; fcstyear0 = fcstyear; }
                //MessageBox.Show("fcst year =" + fcstyear);
                //MessageBox.Show("fcst year0 =" + fcstyear0);
                //MessageBox.Show("thisweeknum =" + thisweeknum);
                //MessageBox.Show("lastweeknum =" + lastweeknum);
            }

            //MessageBox.Show("This week and year: " + thisweeknum + "&" + fcstyear);
            //MessageBox.Show("Last week and year: " + lastweeknum + "&" + fcstyear0);

            fcstthisweek = pro.report_CustomerFcstByWeek(thisweeknum, fcstyear);
            fcstlastweek = pro.report_CustomerFcstByWeek(lastweeknum, fcstyear0);

            DataRow dr = fcstthisweek.NewRow();
            int colnum = fcstthisweek.Columns.Count;
            int rownum = fcstthisweek.Rows.Count;
            int sum = 0;
            for (int i = 2; i < colnum; i++)
            {
                for (int j = 0; j < rownum; j++)
                {
                    sum += int.Parse(fcstthisweek.Rows[j][fcstthisweek.Columns[i]].ToString());
                }
                dr[i] = sum;
                sum = 0;
            }
            dr[1] = "TOTAL";

            fcstthisweek.Rows.Add(dr);

            DataRow drlast = fcstlastweek.NewRow();
            int colnumlast = fcstlastweek.Columns.Count;
            int rownumlast = fcstlastweek.Rows.Count;
            sum = 0;
            for (int i = 2; i < colnumlast; i++)
            {
                for (int j = 0; j < rownumlast; j++)
                {
                    sum += int.Parse(fcstlastweek.Rows[j][fcstlastweek.Columns[i]].ToString());
                }
                drlast[i] = sum;
                sum = 0;
                //MessageBox.Show("SUM = " + sum);
            }
            drlast[1] = "TOTAL";
            fcstlastweek.Rows.Add(drlast);

            DataRow drdiff = fcstthisweek.NewRow();
            int colnumdiff = fcstthisweek.Columns.Count;
            int rownumdiff = fcstthisweek.Rows.Count;
            int diff = 0;
            for (int i = 2; i < colnum - 1; i++)
            {
                diff = int.Parse(fcstthisweek.Rows[rownumdiff - 1][fcstthisweek.Columns[i]].ToString()) - int.Parse(fcstlastweek.Rows[rownumdiff - 1][fcstlastweek.Columns[i + 1]].ToString());
                drdiff[i] = diff;
            }
            drdiff[1] = "DIFFERENCE";
            fcstthisweek.Rows.Add(drdiff);


            DataRow drptram = fcstthisweek.NewRow();
            int colnumptram = fcstthisweek.Columns.Count;
            int rownumptram = fcstthisweek.Rows.Count;
            double ptram = 0.0;
            for (int i = 2; i < colnum - 1; i++)
            {
                int fcsttw = int.Parse(fcstthisweek.Rows[rownumdiff - 1][fcstthisweek.Columns[i]].ToString());
                int fcstlw = int.Parse(fcstlastweek.Rows[rownumdiff - 1][fcstlastweek.Columns[i + 1]].ToString());

                if (fcstlw != 0)
                {
                    ptram = (fcsttw - fcstlw) * 1.0 / fcstlw * 100.0;
                }
                else
                {
                    if (fcsttw == 0)
                    {
                        ptram = 0;
                    }
                    else
                    {
                        ptram = 100;
                    }

                }
                drptram[i] = ptram;
            }

            drptram[1] = "RATE";
            fcstthisweek.Rows.Add(drptram);

            DataRow drW = fcstthisweek.NewRow();
            drW[1] = "WEEK";
            drW[2] = thisweeknum + 1;
            int weeknumcolumnum = fcstthisweek.Columns.Count;

            if (weeknumcolumnum >= 3)
            {
                for (int i = 3; i < weeknumcolumnum; i++)
                {
                    if (int.Parse(drW[i - 1].ToString()) == GetLastWeekNumber(DateTime.Now.Year.ToString() + "-12-31"))
                    {
                        drW[i] = 1;
                    }
                    else
                    {
                        drW[i] = int.Parse(drW[i - 1].ToString()) + 1;
                    }

                }

            }

            fcstthisweek.Rows.InsertAt(drW, 0);



            DataRow drLW = fcstlastweek.NewRow();
            drLW[1] = "WEEK";
            //drLW[2] = (lastweeknum+1 > lastfcstweeknum) ? lastfcstweeknum : lastweeknum + 1;
            drLW[2] = (int.Parse(drW[2].ToString()) - 1) == 0 ? lastfcstweeknum : (int.Parse(drW[2].ToString()) - 1);
            int weeknumcolumnumlast = fcstlastweek.Columns.Count;
            //DateTime tempdate = dateTimePicker1.Value;
            if (weeknumcolumnumlast >= 3)
            {
                for (int i = 3; i < weeknumcolumnumlast; i++)
                {
                    if (int.Parse(drLW[i - 1].ToString()) == lastfcstweeknum)
                    {
                        drLW[i] = 1;
                    }
                    else
                    {
                        drLW[i] = int.Parse(drLW[i - 1].ToString()) + 1;
                    }
                }
            }
            fcstlastweek.Rows.InsertAt(drLW, 0);
            fcstthisweek.Merge(fcstlastweek);

            /*
            for(int i=2;i<fcstthisweek.Columns.Count;i++)
            {
                fcstthisweek.Columns[i].DataType = typeof(double);
            }
            */

            
            backgroundWorker1.ReportProgress(7);
           
            loadPOBalance();
            backgroundWorker1.ReportProgress(8);

        }

        public DataTable dt = new DataTable();
        public DataTable dtPoBalance = new DataTable();
        public DataTable dtPoBalanceType = new DataTable();
        public DataTable dtPoBalanceTypeSS = new DataTable();

        public DataTable dtPo = new DataTable();
        public DataTable dtPoType = new DataTable();
        public DataTable dtPoTypeSS = new DataTable();


        public DataTable fcstthisweek = new DataTable();
        public DataTable fcstlastweek = new DataTable();
        public DataTable dt9 = new DataTable();




        public void initdata2()
        {
            DateTime tempdate = dateTimePicker1.Value;
            ProductBLL pro = new ProductBLL();
           
            string startdate, enddate;
            DateTime FIX_ST_DATE = FirstDayOfWeek(dateTimePicker1.Value);
            DateTime FIX_EN_DATE = LastDayOfWeek(dateTimePicker1.Value);

            DateTime ST_DATE = FIX_ST_DATE;
            DateTime EN_DATE = FIX_EN_DATE;

            startdate = new Form1().STYMD2(ST_DATE.Year, ST_DATE.Month, ST_DATE.Day);
            enddate = new Form1().STYMD2(EN_DATE.Year, EN_DATE.Month, EN_DATE.Day);
            //MessageBox.Show("Start Date =  " + startdate);
            //MessageBox.Show("End Date =  " + enddate);

            //dt = pro.report_WeeklyPOByCustomer();

            for (int i = 0; i <= 10; i++)
            {

                dt = pro.report_WeeklyPOByCustomer2(startdate, enddate);
                dtPo.Merge(dt);
                ST_DATE = ST_DATE.AddDays(-7);
                EN_DATE = EN_DATE.AddDays(-7);
                startdate = new Form1().STYMD2(ST_DATE.Year, ST_DATE.Month, ST_DATE.Day);
                enddate = new Form1().STYMD2(EN_DATE.Year, EN_DATE.Month, EN_DATE.Day);
            }


            dataGridView1.DataSource = dtPo;


            ST_DATE = FIX_ST_DATE;
            EN_DATE = FIX_EN_DATE;
            startdate = new Form1().STYMD2(ST_DATE.Year, ST_DATE.Month, ST_DATE.Day);
            enddate = new Form1().STYMD2(EN_DATE.Year, EN_DATE.Month, EN_DATE.Day);

            for (int i = 0; i <= 10; i++)
            {
                if (checkBox1.Checked == true)
                {
                    dt = pro.report_WeeklyPOByTypeALL2SS(startdate, enddate);
                }
                else
                {
                    dt = pro.report_WeeklyPOByTypeALL2(startdate, enddate);
                }

                dtPoType.Merge(dt);
                ST_DATE = ST_DATE.AddDays(-7);
                EN_DATE = EN_DATE.AddDays(-7);
                startdate = new Form1().STYMD2(ST_DATE.Year, ST_DATE.Month, ST_DATE.Day);
                enddate = new Form1().STYMD2(EN_DATE.Year, EN_DATE.Month, EN_DATE.Day);
            }

            dataGridView5.DataSource = dtPoType;



            ST_DATE = FIX_ST_DATE;
            EN_DATE = FIX_EN_DATE;
            startdate = new Form1().STYMD2(ST_DATE.Year, ST_DATE.Month, ST_DATE.Day);
            enddate = new Form1().STYMD2(EN_DATE.Year, EN_DATE.Month, EN_DATE.Day);

            for (int i = 0; i <= 10; i++)
            {
                dt = pro.report_WeeklyPOByTypeALL2SS(startdate, enddate);
                dtPoTypeSS.Merge(dt);
                ST_DATE = ST_DATE.AddDays(-7);
                EN_DATE = EN_DATE.AddDays(-7);
                startdate = new Form1().STYMD2(ST_DATE.Year, ST_DATE.Month, ST_DATE.Day);
                enddate = new Form1().STYMD2(EN_DATE.Year, EN_DATE.Month, EN_DATE.Day);
            }

            dataGridView7.DataSource = dtPoTypeSS;






            //dt = pro.report_WeeklyPOByType(new Form1().STYMD2(tempdate.Year, tempdate.Month, tempdate.Day));
            // dt = pro.report_WeeklyPOByType(new Form1().STYMD2(tempdate.Year, tempdate.Month, tempdate.Day));

            //dataGridView5.DataSource = dt;


            dt = pro.report_POBalanceByCustomer(new Form1().STYMD2(tempdate.Year, tempdate.Month, tempdate.Day));
            dtPoBalance.Merge(dt);
            tempdate = FridayofWeek(dateTimePicker1.Value.AddDays(-7));
            for (int i = 0; i <= 10; i++)
            {
                dt = pro.report_POBalanceByCustomer(new Form1().STYMD2(tempdate.Year, tempdate.Month, tempdate.Day));
                dtPoBalance.Merge(dt);
                tempdate = tempdate.AddDays(-7);
            }
            dataGridView2.DataSource = dtPoBalance;





            tempdate = dateTimePicker1.Value;
            dt = pro.report_POBalanceByType(new Form1().STYMD2(tempdate.Year, tempdate.Month, tempdate.Day));
            dtPoBalanceType.Merge(dt);
            tempdate = FridayofWeek(dateTimePicker1.Value.AddDays(-7));
            for (int i = 0; i <= 10; i++)
            {
                dt = pro.report_POBalanceByType(new Form1().STYMD2(tempdate.Year, tempdate.Month, tempdate.Day));
                dtPoBalanceType.Merge(dt);
                tempdate = tempdate.AddDays(-7);
            }

            dataGridView6.DataSource = dtPoBalanceType;

            tempdate = dateTimePicker1.Value;
            dt = pro.report_POBalanceByTypeSS(new Form1().STYMD2(tempdate.Year, tempdate.Month, tempdate.Day));
            dtPoBalanceTypeSS.Merge(dt);
            tempdate = FridayofWeek(dateTimePicker1.Value.AddDays(-7));
            for (int i = 0; i <= 10; i++)
            {
                dt = pro.report_POBalanceByTypeSS(new Form1().STYMD2(tempdate.Year, tempdate.Month, tempdate.Day));
                dtPoBalanceTypeSS.Merge(dt);
                tempdate = tempdate.AddDays(-7);
            }
            dataGridView8.DataSource = dtPoBalanceTypeSS;




            int thisweeknum = 0;
            int fcstyear = 0;
            int fcstyear0 = 0;
            int lastfcstweeknum = 0;
            int lastweeknum = 0;



            if (textBox1.Text == "" || textBox2.Text == "")
            {
                thisweeknum = new Form1().GetWeekNumber(dateTimePicker1.Value);
                fcstyear = dateTimePicker1.Value.Year;
                fcstyear0 = fcstyear;
                lastfcstweeknum = pro.getlastFCSTWeekNum(fcstyear);
                // MessageBox.Show("Last FCSt weeknum = " + lastfcstweeknum);

                if (lastfcstweeknum == -1) { fcstyear = fcstyear - 1; fcstyear0 = fcstyear; thisweeknum = 53; lastfcstweeknum = 52; }
                else if (lastfcstweeknum < thisweeknum) thisweeknum = lastfcstweeknum;

                if (thisweeknum == 1) { lastweeknum = 53; fcstyear0 = fcstyear - 1; } else { lastweeknum = thisweeknum - 1; }
            }
            else
            {
                fcstyear = int.Parse(textBox1.Text);
                thisweeknum = int.Parse(textBox2.Text);
                if (thisweeknum == 1) { fcstyear0 = fcstyear - 1; lastweeknum = pro.getlastFCSTWeekNum(fcstyear0); } else { lastweeknum = thisweeknum - 1; fcstyear0 = fcstyear; }
                //MessageBox.Show("fcst year =" + fcstyear);
                //MessageBox.Show("fcst year0 =" + fcstyear0);
                //MessageBox.Show("thisweeknum =" + thisweeknum);
                //MessageBox.Show("lastweeknum =" + lastweeknum);
            }

            //MessageBox.Show("This week and year: " + thisweeknum + "&" + fcstyear);
            //MessageBox.Show("Last week and year: " + lastweeknum + "&" + fcstyear0);

            fcstthisweek = pro.report_CustomerFcstByWeek(thisweeknum, fcstyear);
            fcstlastweek = pro.report_CustomerFcstByWeek(lastweeknum, fcstyear0);

            DataRow dr = fcstthisweek.NewRow();
            int colnum = fcstthisweek.Columns.Count;
            int rownum = fcstthisweek.Rows.Count;
            int sum = 0;
            for (int i = 2; i < colnum; i++)
            {
                for (int j = 0; j < rownum; j++)
                {
                    sum += int.Parse(fcstthisweek.Rows[j][fcstthisweek.Columns[i]].ToString());
                }
                dr[i] = sum;
                sum = 0;
            }
            dr[1] = "TOTAL";

            fcstthisweek.Rows.Add(dr);

            DataRow drlast = fcstlastweek.NewRow();
            int colnumlast = fcstlastweek.Columns.Count;
            int rownumlast = fcstlastweek.Rows.Count;
            sum = 0;
            for (int i = 2; i < colnumlast; i++)
            {
                for (int j = 0; j < rownumlast; j++)
                {
                    sum += int.Parse(fcstlastweek.Rows[j][fcstlastweek.Columns[i]].ToString());
                }
                drlast[i] = sum;
                sum = 0;
                //MessageBox.Show("SUM = " + sum);
            }
            drlast[1] = "TOTAL";
            fcstlastweek.Rows.Add(drlast);

            DataRow drdiff = fcstthisweek.NewRow();
            int colnumdiff = fcstthisweek.Columns.Count;
            int rownumdiff = fcstthisweek.Rows.Count;
            int diff = 0;
            for (int i = 2; i < colnum - 1; i++)
            {
                diff = int.Parse(fcstthisweek.Rows[rownumdiff - 1][fcstthisweek.Columns[i]].ToString()) - int.Parse(fcstlastweek.Rows[rownumdiff - 1][fcstlastweek.Columns[i + 1]].ToString());
                drdiff[i] = diff;
            }
            drdiff[1] = "DIFFERENCE";
            fcstthisweek.Rows.Add(drdiff);


            DataRow drptram = fcstthisweek.NewRow();
            int colnumptram = fcstthisweek.Columns.Count;
            int rownumptram = fcstthisweek.Rows.Count;
            double ptram = 0.0;
            for (int i = 2; i < colnum - 1; i++)
            {
                int fcsttw = int.Parse(fcstthisweek.Rows[rownumdiff - 1][fcstthisweek.Columns[i]].ToString());
                int fcstlw = int.Parse(fcstlastweek.Rows[rownumdiff - 1][fcstlastweek.Columns[i + 1]].ToString());

                if (fcstlw != 0)
                {
                    ptram = (fcsttw - fcstlw) * 1.0 / fcstlw * 100.0;
                }
                else
                {
                    if (fcsttw == 0)
                    {
                        ptram = 0;
                    }
                    else
                    {
                        ptram = 100;
                    }

                }
                drptram[i] = ptram;
            }

            drptram[1] = "RATE";
            fcstthisweek.Rows.Add(drptram);

            DataRow drW = fcstthisweek.NewRow();
            drW[1] = "WEEK";
            drW[2] = thisweeknum + 1;
            int weeknumcolumnum = fcstthisweek.Columns.Count;

            if (weeknumcolumnum >= 3)
            {
                for (int i = 3; i < weeknumcolumnum; i++)
                {
                    if (int.Parse(drW[i - 1].ToString()) == GetLastWeekNumber(DateTime.Now.Year.ToString() + "-12-31"))
                    {
                        drW[i] = 1;
                    }
                    else
                    {
                        drW[i] = int.Parse(drW[i - 1].ToString()) + 1;
                    }

                }

            }

            fcstthisweek.Rows.InsertAt(drW, 0);



            DataRow drLW = fcstlastweek.NewRow();
            drLW[1] = "WEEK";
            //drLW[2] = (lastweeknum+1 > lastfcstweeknum) ? lastfcstweeknum : lastweeknum + 1;
            drLW[2] = (int.Parse(drW[2].ToString()) - 1) == 0 ? lastfcstweeknum : (int.Parse(drW[2].ToString()) - 1);
            int weeknumcolumnumlast = fcstlastweek.Columns.Count;
            //DateTime tempdate = dateTimePicker1.Value;
            if (weeknumcolumnumlast >= 3)
            {
                for (int i = 3; i < weeknumcolumnumlast; i++)
                {
                    if (int.Parse(drLW[i - 1].ToString()) == lastfcstweeknum)
                    {
                        drLW[i] = 1;
                    }
                    else
                    {
                        drLW[i] = int.Parse(drLW[i - 1].ToString()) + 1;
                    }
                }
            }
            fcstlastweek.Rows.InsertAt(drLW, 0);
            fcstthisweek.Merge(fcstlastweek);

            /*
            for(int i=2;i<fcstthisweek.Columns.Count;i++)
            {
                fcstthisweek.Columns[i].DataType = typeof(double);
            }
            */

            dataGridView3.DataSource = fcstthisweek;
            setRowNumber(dataGridView1);
            setRowNumber(dataGridView2);
            setRowNumber(dataGridView6);
            setRowNumber(dataGridView5);
            setRowNumber(dataGridView3);

            formatWeeklyPOByCustomer(dataGridView1);
            formatWeeklyPOBalanceByCustomer(dataGridView2);

            formatWeeklyPOByType(dataGridView5);
            formatWeeklyPOBalanceByType(dataGridView6);
            formatWeeklyPOByType(dataGridView7);
            formatWeeklyPOBalanceByType(dataGridView8);
            if (dataGridView3.Rows.Count >= 9)
                formatWeeklyFCSTByCustomer(dataGridView3);
            loadPOBalance();

        }



        public void formatWeeklyPOByCustomer(DataGridView dataGridView1)
        {
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;


            dataGridView1.Columns["TOTAL_PO_QTY"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["SEV"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["SEVT"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["SAMSUNG_ASIA"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["OTHERS"].DefaultCellStyle.Format = "#,0";

            dataGridView1.Columns["TOTAL_PO_QTY"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["TOTAL_PO_QTY"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["TOTAL_PO_QTY"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);
        }

        public void formatWeeklyPOBalanceByCustomer(DataGridView dataGridView1)
        {
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;

            dataGridView1.Columns["TOTAL_PO_BALANCE"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["SEV"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["SEVT"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["SAMSUNG_ASIA"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["OTHERS"].DefaultCellStyle.Format = "#,0";

            dataGridView1.Columns["TOTAL_PO_BALANCE"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["TOTAL_PO_BALANCE"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["TOTAL_PO_BALANCE"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

        }

        public void formatWeeklyPOByType(DataGridView dataGridView1)
        {
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;

            dataGridView1.Columns["TOTAL_PO_QTY"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["TSP"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["LABEL"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["UV"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["OLED"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["TAPE"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["RIBBON"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["SPT"].DefaultCellStyle.Format = "#,0";

            dataGridView1.Columns["TOTAL_PO_QTY"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["TOTAL_PO_QTY"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["TOTAL_PO_QTY"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

        }

        public void formatWeeklyPOBalanceByType(DataGridView dataGridView1)
        {
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;

            dataGridView1.Columns["TOTAL_PO_BALANCE"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["TSP"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["LABEL"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["UV"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["OLED"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["TAPE"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["RIBBON"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["SPT"].DefaultCellStyle.Format = "#,0";

            dataGridView1.Columns["TOTAL_PO_BALANCE"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["TOTAL_PO_BALANCE"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["TOTAL_PO_BALANCE"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

        }


        public void formatWeeklyFCSTByCustomer(DataGridView dataGridView1)
        {
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;

            dataGridView1.Columns["W1"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W2"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W3"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W4"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W5"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W6"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W7"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W8"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W9"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W10"].DefaultCellStyle.Format = "#,0";
            dataGridView1.Columns["W11"].DefaultCellStyle.Format = "#,0";

            dataGridView1.Rows[6].DefaultCellStyle.Format = "#,0 \\%";

            dataGridView1.Rows[4].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Rows[4].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Rows[4].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);


            dataGridView1.Rows[5].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Rows[5].DefaultCellStyle.BackColor = Color.DarkGreen;
            dataGridView1.Rows[5].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Rows[6].DefaultCellStyle.ForeColor = Color.Black;
            dataGridView1.Rows[6].DefaultCellStyle.BackColor = Color.DarkOrange;
            dataGridView1.Rows[6].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);



            dataGridView1.Rows[11].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Rows[11].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Rows[11].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);


            dataGridView1.Rows[0].DefaultCellStyle.ForeColor = Color.Black;
            dataGridView1.Rows[0].DefaultCellStyle.BackColor = Color.GreenYellow;
            dataGridView1.Rows[0].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Rows[7].DefaultCellStyle.ForeColor = Color.Black;
            dataGridView1.Rows[7].DefaultCellStyle.BackColor = Color.GreenYellow;
            dataGridView1.Rows[7].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);



        }




        public void formatDataGridViewtraPO(DataGridView dataGridView1)
        {
            try
            {
                dataGridView1.EnableHeadersVisualStyles = false;
                dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
                dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.GreenYellow;
                dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("tahoma", 10, FontStyle.Bold);
                dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
                dataGridView1.DefaultCellStyle.SelectionBackColor = Color.Yellow;

                dataGridView1.Columns["TOTAL_PO_QTY"].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns["SEV"].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns["SEVT"].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns["SAMSUNG_ASIA"].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns["OTHERS"].DefaultCellStyle.Format = "#,0";

                dataGridView1.Columns["TOTAL_PO_BALANCE"].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns["TSP"].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns["LABEL"].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns["UV"].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns["OLED"].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns["TAPE"].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns["RIBBON"].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns["SPT"].DefaultCellStyle.Format = "#,0";

                dataGridView1.Columns["W1"].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns["W2"].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns["W3"].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns["W4"].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns["W5"].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns["W6"].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns["W7"].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns["W8"].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns["W9"].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns["W10"].DefaultCellStyle.Format = "#,0";
                dataGridView1.Columns["W11"].DefaultCellStyle.Format = "#,0";
            }
            catch(Exception ex)
            {

            }
           

            /*
            dataGridView1.Columns["PO_QTY"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["PO_QTY"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["PO_QTY"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["PO_BALANCE"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["PO_BALANCE"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["PO_BALANCE"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["TOTAL_DELIVERED"].DefaultCellStyle.ForeColor = Color.White;
            dataGridView1.Columns["TOTAL_DELIVERED"].DefaultCellStyle.BackColor = Color.Gray;
            dataGridView1.Columns["TOTAL_DELIVERED"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["DELIVERED_AMOUNT"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["DELIVERED_AMOUNT"].DefaultCellStyle.BackColor = Color.Black;
            dataGridView1.Columns["DELIVERED_AMOUNT"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["BALANCE_AMOUNT"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["BALANCE_AMOUNT"].DefaultCellStyle.BackColor = Color.Black;
            dataGridView1.Columns["BALANCE_AMOUNT"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);

            dataGridView1.Columns["PO_AMOUNT"].DefaultCellStyle.ForeColor = Color.Yellow;
            dataGridView1.Columns["PO_AMOUNT"].DefaultCellStyle.BackColor = Color.Black;
            dataGridView1.Columns["PO_AMOUNT"].DefaultCellStyle.Font = new Font("tahoma", 9, FontStyle.Bold);


            dataGridView1.Columns["CUST_NAME_KD"].DefaultCellStyle.BackColor = Color.Aqua;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.BackColor = Color.Brown;
            dataGridView1.Columns["G_NAME"].DefaultCellStyle.ForeColor = Color.Yellow;

            */



        }

        public void loadPOBalance()
        {
            ProductBLL pro = new ProductBLL();
            dt9 = pro.report_CustomerPOBalanceByType();
            //dataGridView9.DataSource = dt;
        }



        private void button1_Click(object sender, EventArgs e)
        {

            if (!backgroundWorker1.IsBusy)
            {
                pictureBox1.Show();
                backgroundWorker1.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Đang trong tiến trình khác, thử lại sau");
            }

            /*
            try
            {
                initdata();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Tuần đã chọn chưa nhập FCST\n" + ex.ToString());
            }      
            */
            
        }

        private void reportForm_Load(object sender, EventArgs e)
        {
            //initdata();
            this.ContextMenuStrip = contextMenuStrip1;
            this.dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            this.dataGridView2.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            this.dataGridView5.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            this.dataGridView6.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            this.dataGridView3.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            textBox1.Text = DateTime.Today.Year.ToString();
            textBox2.Text = (new Form1().GetWeekNumber(DateTime.Today)-1).ToString();
            pictureBox1.Hide();
        }
        
        //export bao cao
        private void button2_Click(object sender, EventArgs e)
        {            
            ExcelFactory.exportReportToExcel(dataGridView1, dataGridView2, dataGridView5, dataGridView6, dataGridView3,dataGridView7,dataGridView8,dataGridView9);
        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void save1ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new Form1().CopyToClipboardWithHeaders(dataGridView1);
        }

        private void save2ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new Form1().CopyToClipboardWithHeaders(dataGridView2);
        }

        private void save3ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new Form1().CopyToClipboardWithHeaders(dataGridView5);
        }

        private void save4ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new Form1().CopyToClipboardWithHeaders(dataGridView6);
        }

        private void save5ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new Form1().CopyToClipboardWithHeaders(dataGridView3);
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }
        private void setRowNumber(DataGridView dgv)
        {
            foreach (DataGridViewRow row in dgv.Rows)
            {
                row.HeaderCell.Value = (row.Index + 1).ToString();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            /*
            string Dir = System.IO.Directory.GetCurrentDirectory();           
            string file = Dir + "\\reporttemplate.xlsx";
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dt = pro.traFCSTYERKWEEK("2021", "2");
            ExcelFactory.fastexport(dt, file);
            MessageBox.Show("Export Bao Cao Hoan Thanh !");
            */

            string Dir = System.IO.Directory.GetCurrentDirectory();
            string file = Dir + "\\testreport.xlsx";
            MessageBox.Show(file);
            ProductBLL pro = new ProductBLL();
            DataTable dt = new DataTable();
            dt = pro.traFCSTYERKWEEK("2020","2");
            string savepath = Dir + "\\REPORT-HUNG.xlsx";
            ExcelFactory.editFileExcelReport(file, dt, savepath);
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                initdata3();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Tuần đã chọn chưa nhập FCST\n" + ex.ToString());
            }
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            label12.Text = e.ProgressPercentage + "/8";
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            pictureBox1.Hide();
            dataGridView1.DataSource = dtPo;
            dataGridView5.DataSource = dtPoType;
            dataGridView7.DataSource = dtPoTypeSS;
            dataGridView2.DataSource = dtPoBalance;
            dataGridView6.DataSource = dtPoBalanceType;
            dataGridView8.DataSource = dtPoBalanceTypeSS;
            dataGridView3.DataSource = fcstthisweek;
            dataGridView9.DataSource = dt9;

            setRowNumber(dataGridView1);
            setRowNumber(dataGridView2);
            setRowNumber(dataGridView6);
            setRowNumber(dataGridView5);
            setRowNumber(dataGridView3);

            formatWeeklyPOByCustomer(dataGridView1);
            formatWeeklyPOBalanceByCustomer(dataGridView2);

            formatWeeklyPOByType(dataGridView5);
            formatWeeklyPOBalanceByType(dataGridView6);
            formatWeeklyPOByType(dataGridView7);
            formatWeeklyPOBalanceByType(dataGridView8);
            if (dataGridView3.Rows.Count >= 9)
                formatWeeklyFCSTByCustomer(dataGridView3);

        }
    }
}
