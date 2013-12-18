using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Printing;
using System.Management;
using System.ServiceProcess;

namespace SpoolCleaner
{
    public partial class Form1 : Form
    {
        static System.Windows.Forms.Timer myTimer = new System.Windows.Forms.Timer();

        public Form1()
        {
            InitializeComponent();


            //here is how to call the methods
            //http://msdn.microsoft.com/en-us/library/cc146163.aspx
            //SetPowerState would be a good method to call (http://msdn.microsoft.com/en-us/library/aa393485(v=vs.85).aspx)
            //   but it has some glaring problems such as not being implemented apparently so yeah theres that.
            //Mayhaps we can ghetto implement this.  I should discuss this with Dan.
            ManagementScope Scope = new ManagementScope(String.Format("\\\\{0}\\root\\CIMV2", "localhost"), null);
            Scope.Connect();
            ObjectQuery Query = new ObjectQuery("SELECT * FROM Win32_Printer");
            ManagementObjectSearcher Searcher = new ManagementObjectSearcher(Scope, Query);

            //foreach (ManagementObject WmiObject in Searcher.Get())
            //{
            //    Console.WriteLine(WmiObject.GetText(TextFormat.CimDtd20));
            //}

            myTimer.Tick += new EventHandler(TimerEventProcessor);

            myTimer.Interval = 5000;
            myTimer.Start();
         
        }
        
        // This is the method to run when the timer is raised. 
        private static void TimerEventProcessor(Object myObject, EventArgs myEventArgs)
        {
            try
            {
                PrintServer ps = new PrintServer(PrintSystemDesiredAccess.None);

                PrintQueueCollection pqc = ps.GetPrintQueues();
                foreach (PrintQueue pq in pqc)
                {
                    pq.Refresh();
                    if (pq.NumberOfJobs != 0)
                    {
                        PrintJobInfoCollection pjic = pq.GetPrintJobInfoCollection();
                        foreach (PrintSystemJobInfo v in pjic)
                        {
                            DateTime dt = DateTime.UtcNow;
                            TimeSpan ts = (dt - v.TimeJobSubmitted);
                            int ms = (int) ts.TotalMilliseconds;
                            if (ms > 7000)
                            {
                                ServiceController sc = new ServiceController("Spooler");

                                if (sc.CanStop)
                                {
                                    sc.Stop();
                                }

                                while (sc.Status == ServiceControllerStatus.Running)
                                {
                                    sc.Refresh();
                                }

                                sc.Start();
                            }
                        }
                    }
                }
            }
            catch
            {
            }
        }
    }
}
