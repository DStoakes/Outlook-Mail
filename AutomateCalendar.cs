using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OutLook = Microsoft.Office.Interop.Outlook;
using Microsoft.VisualBasic;
using System.Windows.Forms;


namespace OutlookMail
{

    //****************************************************
    // This class is designed to automate a users outlook
    // calendar.
    //****************************************************

    class AutomateCalendar : MasterClass
    {
        public AutomateCalendar()
        {
            
            //****************************************************
            // Itterate through "inbox" sub folders to Find 
            // "Calendar" Folder
            //****************************************************

            foreach (OutLook.MAPIFolder inbox in myInbox.Folders)
            {
                if (inbox.Name == "Calendar")
                {
                    Calendar = inbox;
                    OutLook.Items items = Calendar.Items;

                    //****************************************************
                    // Itterate through item in the "Calendar" Folder,
                    // checks each item fits certain criteria.
                    //****************************************************

                    foreach (OutLook.MailItem mail in items)
                    {
                        if (mail.Subject.Contains("Booking") || mail.Subject.Contains("booking") || mail.Subject.Contains("Appointment Confirmation") || mail.Subject.Contains("2991"))
                        {

                            //****************************************************
                            // If the criteria is met, the class then goes through
                            // body of the email to find the company job number
                            // and appointment dates.
                            //****************************************************

                            string subject = mail.Subject.Remove(0, mail.Subject.IndexOf("2991"));
                            JNumber = subject.Remove(8, subject.Length - 8);
                            string emailBody = mail.Body.Remove(0, mail.Body.IndexOf("Your full day appointment is scheduled for "));
                            string emailBodyRemoved = emailBody.Remove(0, 43);
                            int count = emailBodyRemoved.IndexOf("from");
                            string emailDateTime = emailBodyRemoved.Remove(emailBodyRemoved.IndexOf("from"), emailBodyRemoved.Length - count);

                            string[] words = emailDateTime.Split(' ');

                            string day = words[1].Remove(words[1].Length - 2);

                            string month = words[2].Substring(0, 3);

                            string year = words[3];

                            string date = day + "/" + month + "/" + year;


                            try
                            {
                                time = DateTime.Parse(date);
                            }
                            catch
                            {
                                dialogResult = MessageBox.Show("Please Check Date Is Correct: " + JNumber + " " + time, "Check Date", MessageBoxButtons.YesNo);
                                if (dialogResult == DialogResult.No)
                                {
                                    string input = Microsoft.VisualBasic.Interaction.InputBox("Please Enter Correct Date: ", "Check Date");
                                    time = DateTime.Parse(input);
                                }
                            }

                            //****************************************************
                            // Once this information has been collected the class
                            // then creates a new outlook appointment item and
                            // checks if there is already an existing appointment
                            // in the outlook calendar.
                            //****************************************************

                            newAppointment = (OutLook.AppointmentItem)outlookApp.CreateItem(OutLook.OlItemType.olAppointmentItem);
                            newAppointment.Start = time;

                            int CalendarItemCount = 0;
                            if (mail.Subject.Contains("2991"))
                            {
                                newAppointment.Subject = JNumber;

                                foreach (OutLook.AppointmentItem CalendarItem in CalendarFolder.Items)
                                {

                                    //****************************************************
                                    // If an appointment already exists the user will be 
                                    // given the opportunity to check if the email for any
                                    // errors or alternitively override the dates set in
                                    // email body and create new dates.
                                    //****************************************************

                                    if (CalendarItem.Subject == newAppointment.Subject)
                                    {
                                        CalendarItemCount = CalendarItemCount + 1;
                                        Console.WriteLine("Job Already Exists");
                                        dialogResult = MessageBox.Show("Please Check Date Is Correct: " + newAppointment.Subject + " " + time, "Check Date", MessageBoxButtons.YesNo);
                                        if (dialogResult == DialogResult.No)
                                        {
                                            string input = Microsoft.VisualBasic.Interaction.InputBox("Please Enter Correct Date: ", "Check Date");
                                            time = DateTime.Parse(input);
                                            newAppointment.Start = time;
                                            CalendarItem.Delete();
                                            newAppointment.Save();
                                        }
                                        else if (dialogResult == DialogResult.Yes)
                                        {
                                            CalendarItem.Delete();
                                            newAppointment.Save();
                                        }
                                    }
                                }

                                //****************************************************
                                // If the appointment does not exist the class will
                                // then save the appointment to the outlook calendar
                                // and then archive the email to the correct folder.
                                //****************************************************

                                if (CalendarItemCount == 0)
                                {
                                    newAppointment.Save();
                                }
                                foreach (OutLook.MAPIFolder inbox2 in myInbox.Folders)
                                {
                                    if (inbox2.Name == "CompanyName")
                                    {
                                        CompanyName = inbox2;
                                        mail.Move(CompanyName);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}