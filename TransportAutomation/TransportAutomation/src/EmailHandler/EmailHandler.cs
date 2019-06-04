using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using System.IO;
using System.Windows.Forms;

namespace TransportAutomation.src.EmailHandler
{
    class EmailHandler
    {
        public string emailAttachmentsPath { get; set; }
        public string DAIRPath { get; set; }
        private string images { get; set; }
        private string misc { get; set; }
        private string datmrMatmr { get; set; }
        private string dailyReport { get; set; }
        private string snowiz { get; set; }
        private string timesheet { get; set; }
        private string vehicleInspection { get; set; }
        private string journal { get; set; }



        public EmailHandler ()
        {
            this.emailAttachmentsPath = Directory.GetCurrentDirectory() + "\\Email Attachments";
            this.DAIRPath = this.emailAttachmentsPath + "\\DAIR";
            this.images = this.emailAttachmentsPath + "\\Image";
            this.misc = this.emailAttachmentsPath + "\\Unsorted - Sort Manually";
            this.datmrMatmr = this.emailAttachmentsPath + "\\DATMR MATMR";
            this.dailyReport = this.emailAttachmentsPath + "\\Daily Report";
            this.snowiz = this.emailAttachmentsPath + "\\SNOWIZ";
            this.timesheet = this.emailAttachmentsPath + "\\Timesheet";
            this.vehicleInspection = this.emailAttachmentsPath + "\\Vehicle Inspection";
            this.journal = this.emailAttachmentsPath + "\\Journal";
        }
        public MAPIFolder findMailFolder()
        {
            Microsoft.Office.Interop.Outlook.Application app = new Microsoft.Office.Interop.Outlook.Application();
            Accounts accounts = app.Session.Accounts;
            _NameSpace ns = app.GetNamespace("MAPI");

            Folders folders = ns.Folders;
            MAPIFolder reportsFolder = null;

            foreach (MAPIFolder f in folders)
            {
                if (f.Name == /*"Transport OPS2"*/ "transport@krg.ca")
                {
                    reportsFolder = f;
                }
            }
            return reportsFolder;
        }

        public void EnumerateFolders(MAPIFolder folder, bool getAllAttachments)
        {
            Folders childFolders = folder.Folders;
            if (childFolders.Count > 0)
            {
                foreach (Folder childFolder in childFolders)
                {
                    // We only want Inbox folders - ignore Contacts and others
                    if (childFolder.FolderPath.Contains("Inbox"))
                    {
                        IterateMessages(childFolder, getAllAttachments);
                        // Call EnumerateFolders using childFolder, to see if there are any sub-folders within this one
                        EnumerateFolders(childFolder, getAllAttachments);
                    }
                }
            }
        }

        public void IterateMessages(MAPIFolder folder, bool readAll)
        {
            Directory.CreateDirectory(emailAttachmentsPath);
            Directory.CreateDirectory(DAIRPath);
            Directory.CreateDirectory(images);
            Directory.CreateDirectory(misc);
            Directory.CreateDirectory(datmrMatmr);
            Directory.CreateDirectory(dailyReport);
            Directory.CreateDirectory(snowiz);
            Directory.CreateDirectory(vehicleInspection);
            Directory.CreateDirectory(journal);
            Directory.CreateDirectory(timesheet);

            var fi = folder.Items;
            int emailCounter = 0;
            int attachmentsCounter = 0;
            if (fi != null)
            {
                foreach (Object item in fi)
                {
                    
                    MailItem mi = (MailItem)item;
                    string compare;
                    if (readAll)
                    {
                        compare = "hello world";
                    } else
                    {
                        compare = "1";
                    }
                    if (mi.FlagRequest != compare) {
                        emailCounter++;
                        var attachments = mi.Attachments;
                        if (attachments.Count != 0)
                        {
                            for (int i = 1; i <= mi.Attachments.Count; i++)
                            {
                                attachmentsCounter++;
                                string fileName = mi.Attachments[i].FileName;
                                //var date = mi.CreationTime.ToString();
                                //int spaceIndex = date.IndexOf(" ");
                                //date = date.Substring(0, spaceIndex);
                                string sender = mi.SenderName;
                                string savedFileName = sender + " - " + fileName;
                                if (fileName.Contains(".png") || fileName.Contains(".jpg"))
                                {
                                    mi.Attachments[i].SaveAsFile(images + "\\" + savedFileName);
                                } else if (fileName.IndexOf("dair", StringComparison.OrdinalIgnoreCase) >= 0 || fileName.IndexOf("daily airport inspection", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    mi.Attachments[i].SaveAsFile(DAIRPath + "\\" + savedFileName);
                                } else if ((fileName.IndexOf("datmr", StringComparison.OrdinalIgnoreCase) >= 0) || (fileName.IndexOf("matmr", StringComparison.OrdinalIgnoreCase) >= 0))
                                {
                                    mi.Attachments[i].SaveAsFile(datmrMatmr + "\\" + savedFileName);
                                }
                                else if (fileName.IndexOf("daily report", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    mi.Attachments[i].SaveAsFile(dailyReport + "\\" + savedFileName);
                                }
                                else if (fileName.IndexOf("snow", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    mi.Attachments[i].SaveAsFile(snowiz + "\\" + savedFileName);
                                }
                                else if (fileName.IndexOf("vehicle inspection", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    mi.Attachments[i].SaveAsFile(vehicleInspection + "\\" + savedFileName);
                                }
                                else if (fileName.IndexOf("journal", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    mi.Attachments[i].SaveAsFile(journal + "\\" + savedFileName);
                                }
                                else if (fileName.IndexOf("timesheet", StringComparison.OrdinalIgnoreCase) >= 0
                                    || fileName.IndexOf("time sheet", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    mi.Attachments[i].SaveAsFile(timesheet + "\\" + savedFileName);
                                }
                                else
                                {
                                    mi.Attachments[i].SaveAsFile(misc + "\\" + savedFileName);
                                }
                                Directory.CreateDirectory(vehicleInspection);
                                Console.WriteLine("Downloading Attachment: " + fileName);
                            }
                        }
                        // 1 for processed, null or "Follow up" for unprocessed
                            mi.FlagRequest = "1";
                    }
                    mi.Save();
                    
                }
                MessageBox.Show("Processed " + emailCounter + " emails and downloaded " + attachmentsCounter + 
                    " attachments. Unsorted attachments are stored in 'Unsorted - Sort Manually'"); 
            }
        }
    }    
}
