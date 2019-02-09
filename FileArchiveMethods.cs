using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;
using Microsoft.VisualBasic;
using System.Windows.Forms;

namespace OutlookMail
{

    //****************************************************
    // This class is designed to archive different types of
    // quote files and job files to their correct quote
    // folder or job folder.
    //****************************************************

    public class FileArchiveMethods : MasterClass
    {

        //****************************************************
        // Defines FileName
        //****************************************************

        public static void FileNameArray()
        {
            files = Directory.GetFiles(ArchiveFolder);
            arrCount = files.Count();
            if (arrCount > 0)
            {
                foreach (string file in files)
                {
                    FileName = Path.GetFileName(file);
                }
            }
        }

        //****************************************************
        // Determins if the filename is a quote starting with
        // "Y" and the length of the filename and then moves 
        // the file to its correct folder.
        //**************************************************** 

        public static void YQQuotes()
        {
            foreach (string file in files)
            {
                filearray = file;
                FileName = Path.GetFileName(file);

                if (FileName == null)
                {
                    Console.WriteLine("NULL");
                }

                else if (FileName.Contains("Y"))
                {
                    string FileNameReduced = FileName.Remove(0, 1);
                    int Datum = FileNameReduced.IndexOf("-");
                    QuoteFolderLink2 = @"\" + FileNameReduced.Remove(Datum, FileName.Length - (Datum + 1));

                    if (Datum == 6)
                    {
                        QuoteFolderLink1 = FileNameReduced.Remove(3, FileNameReduced.Length - 3) + "XXX";
                        DestinationFolder = QuoteFolder + QuoteFolderLink1 + QuoteFolderLink2 + @"\" + FileName;
                        File.Move(file, DestinationFolder);
                        Process.Start(QuoteFolder + QuoteFolderLink1 + QuoteFolderLink2);
                        count = 1;
                    }
                    else if (Datum == 7)
                    {
                        QuoteFolderLink1 = @"\" + FileNameReduced.Remove(4, FileNameReduced.Length - 4) + "XXX";
                        DestinationFolder = QuoteFolder + QuoteFolderLink1 + QuoteFolderLink2 + @"\" + FileName;
                        if (File.Exists(DestinationFolder))
                        {
                            ArchiveFile();
                        }
                        else
                        {
                            File.Move(file, DestinationFolder);
                            Process.Start(QuoteFolder + QuoteFolderLink1 + QuoteFolderLink2);
                            count = 1;

                        }
                    }
                    else
                    {
                        count = 0;
                    }
                }
            }
        }

        //****************************************************
        // Determins if the filename is a quote starting with
        // "Q" and the length of the filename and then moves 
        // the file to its correct folder.
        //****************************************************

        public static void QQuotes()
        {
            foreach (string file in files)
            {
                FileName = Path.GetFileName(file);

                if (FileName == null)
                {
                    Console.WriteLine("NULL");
                }

                else if (!FileName.Contains("Y") && FileName.Contains("Q"))
                {
                    Console.WriteLine("Q");
                    int Datum = FileName.IndexOf("-");
                    QuoteFolderLink2 = @"\" + FileName.Remove(Datum, FileName.Length - (Datum));
                    if (Datum == 6)
                    {
                        QuoteFolderLink1 = @"\" + FileName.Remove(3, FileName.Length - 3) + "XXX";
                        DestinationFolder = QuoteFolder + QuoteFolderLink1 + QuoteFolderLink2 + @"\" + FileName;
                        File.Move(file, DestinationFolder);
                        Process.Start(QuoteFolder + QuoteFolderLink1 + QuoteFolderLink2);
                        count = 1;
                    }
                    else if (Datum == 7)
                    {
                        QuoteFolderLink1 = @"\" + FileName.Remove(4, FileName.Length - 4) + "XXX";
                        DestinationFolder = QuoteFolder + QuoteFolderLink1 + QuoteFolderLink2 + @"\" + FileName;
                        File.Move(file, DestinationFolder);
                        Process.Start(QuoteFolder + QuoteFolderLink1 + QuoteFolderLink2);
                        count = 1;
                    }
                    else
                    {
                        count = 0;
                    }
                }
            }
        }

        //****************************************************
        // Determins the Length of the filename and if it
        // strictly has numbers only, indicating a "job number"
        // This section is not yet finished, next steps would
        // involve moving the file to its correct job folder.
        //****************************************************

        public static void JobNumber()
        {
            foreach (string file in files)
            {
                FileName = Path.GetFileName(file);

                if (FileName == null)
                {
                    Console.WriteLine("NULL");
                }
                else
                {
                    QuoteFolderLink1 = FileName.Remove(5, FileName.Length - 5);
                    int number;
                    if (int.TryParse(QuoteFolderLink1, out number) == true)
                    {
                        count = 1;
                    }
                    else
                    {
                        count = 0;
                    }
                }
            }
        }

        //****************************************************
        // This method displays an error message to the console
        // if the filename is not recognised.
        //****************************************************

        public static void Error()
        {
            if (count == 0)
            {
                Console.WriteLine("Error " + FileName + " File Name Not Recognized");
            }
        }

        //****************************************************
        // This method clears the files array so it does not
        // rename files incorrectly when the program is repeated.
        //****************************************************

        public static void ClearArray()
        {
            Array.Clear(files, 0, files.Length);
        }

        //****************************************************
        // This method handles files if the file already exists
        // in the quote folder, this method presents a dialog box
        // allowing the user to rename the file or archive the
        // existing file if it is being supersedded.
        //****************************************************

        public static void ArchiveFile()
        {
            Process.Start(QuoteFolder + QuoteFolderLink1 + QuoteFolderLink2);
            DialogResult dialogResult = MessageBox.Show("Quote Already Exists, Would You Like To Archive The Existing Quote??", "Quote Exists", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                string archiveFolder = QuoteFolder + QuoteFolderLink1 + QuoteFolderLink2 + @"\8.AdditionalInfo\Archive\" + FileName + "(OLDV" + ArchiveCount + ")";
                string existingFile = QuoteFolder + QuoteFolderLink1 + QuoteFolderLink2 + @"\" + FileName;
                if (File.Exists(archiveFolder))
                {
                    while (File.Exists(archiveFolder))
                    {
                        Console.WriteLine(archiveFolder.Remove(archiveFolder.IndexOf("OLDV"), 5));
                        count = 1;
                    }
                }
                else
                {

                    File.Move(existingFile, archiveFolder);
                    File.Move(filearray, DestinationFolder);
                    count = 1;
                }

            }
            else if (dialogResult == DialogResult.No)
            {
                string input = Microsoft.VisualBasic.Interaction.InputBox("Please Rename Quote");
                string inputDestinationFolder = QuoteFolder + QuoteFolderLink1 + QuoteFolderLink2 + @"\" + input + ".pdf";
                File.Move(filearray, inputDestinationFolder);
                count = 1;
            }
        }
    }
}