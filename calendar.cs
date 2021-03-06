﻿using System;
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

        public Form1()
        {
            InitializeComponent();
        }

        void addAppointment(int tyear, int tmonth, int tday, int tstarttime, int tendtime, string tactivity, string tlocation)
        {
            appointments.Add(new appointment( tyear, tmonth, tday, tstarttime, tendtime, tactivity, tlocation));
        }


        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            Excel.Range oRng;

            oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = true;

            //Get a new workbook.
            oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
            oSheet = (Excel._Worksheet)oWB.ActiveSheet;

            //Add table headers going cell by cell.
            oSheet.Cells[1, 1] = "Date";
            oSheet.Cells[1, 2] = "Start Time";
            oSheet.Cells[1, 3] = "End Time";
            oSheet.Cells[1, 4] = "Activity";
            oSheet.Cells[1, 5] = "Location";

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

            oXL.Visible = true;
            oXL.UserControl = true;

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

        public appointment (int tyear, int tmonth, int tday, int tstarttime, int tendtime, string tactivity, string tlocation){
             year = tyear;
             month = tmonth;
             day = tday;
             startTime = tstarttime;
             endTime = tendtime;
             activity = tactivity;
             location = tlocation;
        }

            public string showDate (){
                string monthName;
                switch (month)
                {
                    case(1):
                        monthName = "Jan";
                        break;
                    case(2):
                        monthName = "Feb";
                        break;
                    case(3):
                        monthName = "Mar";
                        break;
                    case(4):
                        monthName = "Apr";
                        break;
                    case(5):
                        monthName = "May";
                        break;
                    case(6):
                        monthName = "Jun";
                        break;
                    case(7):
                        monthName = "Jul";
                        break;
                    case(8):
                        monthName = "Aug";
                        break;
                    case(9):
                        monthName = "Sep";
                        break;
                    case(10):
                        monthName = "Oct";
                        break;
                    case(11):
                        monthName = "Nov";
                        break;
                    case(12):
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
