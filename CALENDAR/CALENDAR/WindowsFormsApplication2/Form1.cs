using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;


namespace WindowsFormsApplication2
{
    public partial class Form1 : Form
    {
        List<appointment> appointments = new List<appointment>();

        void addAppointment(int tyear, int tmonth, int tday, int tstarttime, int tendtime, string tactivity, string tlocation)
        {
            appointments.Add(new appointment(tyear, tmonth, tday, tstarttime, tendtime, tactivity, tlocation));
        }

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            //Excel.Worksheets sheets;
            //Excel.Worksheet oSheet;

            //Excel.Range oRng;//should probably learn how to use ranges, is more efficient.

            oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = true;

            //Get a new workbook.
            //oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
            oWB = oXL.Workbooks.Open("Calendar.xlsx");
            //Excel.Worksheet oSheet = oWB.ActiveSheet;
            oSheet = (Excel._Worksheet)oWB.ActiveSheet;

            //Add table headers going cell by cell.
            oSheet.Cells[1, 1] = "Date";
            oSheet.Cells[1, 2] = "Start Time";
            oSheet.Cells[1, 3] = "End Time";
            oSheet.Cells[1, 4] = "Activity";
            oSheet.Cells[1, 5] = "Location";
            oSheet.Cells.Interior.Color = 255;

            addAppointment(1995, 2, 21, 0, 0, "stuff", "New York");
            addAppointment(2013, 7, 23, 0600, 1200, "orientation", "Stony Brook");
            addAppointment(2055, 2, 21, 1200, 1300, "robot apocalypse", "New Paris");

            for (int i = 0; i < appointments.Count(); i++ )
            {
                oSheet.Cells[i + 2, 1] = appointments[i].showDate();
                oSheet.Cells[i + 2, 2] = appointments[i].startTime;
                oSheet.Cells[i + 2, 3] = appointments[i].endTime;
                oSheet.Cells[i + 2, 4] = appointments[i].activity;
                oSheet.Cells[i + 2, 5] = appointments[i].location;
            }
            //oWB.Save();
            //oWB.Close();
            oXL.Visible = true;
            oXL.UserControl = true;
            ((Excel.Worksheet)oSheet).Change += new Microsoft.Office.Interop.Excel.DocEvents_ChangeEventHandler(anmdRange_Change);

            //Marshal.ReleaseComObject();//don't know what this does
        }
        private void addNamedRanges(Excel.Application oXL)
        {

            //Excel._Workbook oWB;
            //oWB = oXL.Workbooks.Open("Calendar.xlsx");
            //Excel.Range anmdRange;
            //Excel.Worksheet asht;
            //System.Windows.Forms.Control.ControlCollection cc;

            //asht = oWB.ActiveSheet;
            //cc = asht.Controls;
            //anmdRange = cc.AddNamedRange(asht.InnerObject.Range("a1"), "somenamedRange");

            //asht.Change += new Microsoft.Office.Interop.Excel.DocEvents_ChangeEventHandler(anmdRange_Change);

        }

        void anmdRange_Change(Microsoft.Office.Interop.Excel.Range Target)
        {
            MessageBox.Show(Target.Value2.ToString());
        } 
    }


    public class appointment
    {
        public int year;
        public int month;
        public int day;
        public int startTime;
        public int endTime;
        public string activity;
        public string location;

        public appointment(int tyear, int tmonth, int tday, int tstarttime, int tendtime, string tactivity, string tlocation)
        {
            year = tyear;
            month = tmonth;
            day = tday;
            startTime = tstarttime;
            endTime = tendtime;
            activity = tactivity;
            location = tlocation;
        }

        public string showDate()
        {
            string monthName;
            switch (month)
            {
                case (1):
                    monthName = "Jan";
                    break;
                case (2):
                    monthName = "Feb";
                    break;
                case (3):
                    monthName = "Mar";
                    break;
                case (4):
                    monthName = "Apr";
                    break;
                case (5):
                    monthName = "May";
                    break;
                case (6):
                    monthName = "Jun";
                    break;
                case (7):
                    monthName = "Jul";
                    break;
                case (8):
                    monthName = "Aug";
                    break;
                case (9):
                    monthName = "Sep";
                    break;
                case (10):
                    monthName = "Oct";
                    break;
                case (11):
                    monthName = "Nov";
                    break;
                case (12):
                    monthName = "Dec";
                    break;
                default:
                    monthName = "" + month;
                    break;
            }


            return monthName + " " + day + " " + year;
        }




    }
}

