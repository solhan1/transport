using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TransportAutomation.src.DocumentProcessors.docx;
using TransportAutomation.src.EmailHandler;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Office.CustomUI;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using TransportAutomation.src.Logger;

namespace TransportAutomation
{
    class Program
    {
        static void Main(string[] args)
        {
            EmailHandler emailHandler = new EmailHandler();
            DocumentProcessor d = new DocumentProcessor();

            Console.WriteLine("Starting ... \n");

            Console.WriteLine("If this is your first time running this program, you must process all emails (option 1).");
            Console.WriteLine("\nPlease type in an option and press ENTER to proceed.");
            Console.WriteLine("0: Exit.");
            Console.WriteLine("1: Read all emails and download and sort all attachments.");
            Console.WriteLine("2: Read new emails only, and download and sort all attachments.");
            bool readAll = true;
            int option;
            while (true)
            {
                if (int.TryParse(Console.ReadLine(), out option))
                {
                    if (option == 0)
                    {
                        Environment.Exit(0);
                    }
                    else if (option == 1)
                    {
                        break;
                    }
                    else if (option == 2)
                    {
                        Console.WriteLine("Digging through your mailbox ...");
                        readAll = false;
                        break;
                    }
                    else
                    {
                        Console.WriteLine("Invalid number. Please pick a number from 0 to 8.");
                    }
                }
                else
                {
                    Console.WriteLine("Please enter a number.");
                }
            }

            string currentPath = Directory.GetCurrentDirectory();
            string errorPath = currentPath + "\\Failed Documents";
            int year = DateTime.Now.Year;
            int month = DateTime.Now.Month;
            var now = DateTime.Now.ToString();
            now = now.Replace(":", " ");
            string logDirectory = currentPath + "\\Logs" + "\\" + year + "\\" + month;
            string logPath = logDirectory + "\\" + now + ".txt";
            Directory.CreateDirectory(logDirectory);

            try
            {
                MAPIFolder reportsFolder = emailHandler.findMailFolder();

                if (reportsFolder == null)
                {
                    MessageBox.Show("Could not find transport email directory.");
                }
                else
                {
                    Console.WriteLine("Checking for unprocessed emails ..\n");
                    emailHandler.EnumerateFolders(reportsFolder, readAll);
                }

                Console.WriteLine("\nPlease enter the type of file to parse and store in the database. Ensure that all reports are closed.");
                Console.WriteLine("NOTE: Only DAIR's (1) are available for parsing. Storing in the database is under development.");
                Console.WriteLine("\nPlease type in an option and press ENTER to proceed.");
                Console.WriteLine("0: Exit");
                Console.WriteLine("1: DAIR");
                Console.WriteLine("2: Journal");
                Console.WriteLine("3: Image");
                Console.WriteLine("4: DATMR/MATMR");
                Console.WriteLine("5: Daily Report");
                Console.WriteLine("6: SNOWIZ");
                Console.WriteLine("7: Timesheet");
                Console.WriteLine("8: Vehicle Insepection");
                int x;
                while (true)
                {
                    if (int.TryParse(Console.ReadLine(), out x))
                    {
                        if (x == 0)
                        {
                            Environment.Exit(0);
                        }
                        else if (x == 1)
                        {
                            string DAIRPath = emailHandler.DAIRPath;
                            Console.WriteLine("Parsing DAIR's ...\n");
                            d.DAIRparser(DAIRPath);
                            break;
                        }
                        else if (x > 8)
                        {
                            Console.WriteLine("Invalid number. Please pick a number from 0 to 8.");
                        }
                        else
                        {
                            Console.WriteLine("Currently unavailable.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("Please enter a number.");
                    }
                }
            }
            catch (System.Exception e)
            {
                Logger logger = new Logger(logPath);
                logger.Append(e.Message);
            }

            

        }
    }
}
