using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OutLook = Microsoft.Office.Interop.Outlook;
using System.Diagnostics;

//****************************************************
// This class is work in progress, this class would be
// used to collate all the required paperwork for an order
// and merge them together in a PDF format
//****************************************************  

namespace OutlookMail
{
    public class ProcessOrders : MasterClass
    {
        public static void Order()
        {
            string path = @"\\Personal Folders\Inbox\OL Test";
            Process.Start(path);
            Console.ReadKey();
        }
    }
}
