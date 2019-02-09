using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.IO;
using OutLook = Microsoft.Office.Interop.Outlook;
using System.Web;
using System.Net;

namespace OutlookMail
{
    class Program : FileArchiveMethods
    {

        //****************************************************
        // Program sleeps for 1 minute to allow system to boot 
        // then loops through each class.
        //****************************************************

        static void Main(string[] args)
        {
            Thread.Sleep(60000);
            while (true)
            {

                //ProcessOrders.Order(); Work in progress

                AutomateJobs AJ = new AutomateJobs();
                AutomateCalendar AC = new AutomateCalendar();

                FileNameArray();
                if (arrCount > 0)
                {
                    YQQuotes();
                    QQuotes();
                    JobNumber();
                    Error();
                    ClearArray();
                }

                Thread.Sleep(2000);
            }
        }
    }
}