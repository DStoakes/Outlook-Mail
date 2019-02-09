using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OutLook = Microsoft.Office.Interop.Outlook;

namespace OutlookMail
{
    //****************************************************
    // The MasterClass declares reusable objects.
    //****************************************************

    public class MasterClass
    {
        // Outlook Mail Class

        public OutLook.MAPIFolder Trigger;
        public OutLook.MAPIFolder Processed;
        public OutLook.MAPIFolder CompanyName;
        public OutLook.MAPIFolder Calendar;
        public static OutLook.Application outlookApp = new OutLook.Application();
        public static OutLook.NameSpace olNameSpace = outlookApp.GetNamespace("MAPI");
        public static OutLook.MAPIFolder myInbox = olNameSpace.GetDefaultFolder(OutLook.OlDefaultFolders.olFolderInbox);
        public OutLook.MAPIFolder CalendarFolder = olNameSpace.GetDefaultFolder(OutLook.OlDefaultFolders.olFolderCalendar);
        public OutLook.AppointmentItem newAppointment;
        public System.Windows.Forms.DialogResult dialogResult;
        public DateTime time;
        public static string JNumber;

        // FileArchive Master Paths

        public static string ArchiveFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\" + "Job Archiver";
        public static string QuoteFolder = @"X:\Clients\!Quotes";
        public static string[] files;
        public static string FileName;
        public static string QuoteFolderLink1;
        public static string QuoteFolderLink2;
        public static string DestinationFolder;
        public static int count = 0;
        public static int arrCount;
        public static int ArchiveCount = 1;
        public static string filearray;
        public static int class1int = 0;

    }
}
