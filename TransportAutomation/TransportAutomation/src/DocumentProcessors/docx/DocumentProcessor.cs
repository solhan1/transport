using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.IO.Packaging;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Windows.Forms;

namespace TransportAutomation.src.DocumentProcessors.docx
{
    public class DocumentProcessor
    {
        public DocumentProcessor()
        {
        }

        // opens a .docx using a path.
        // filePath: the location of the file
        // readOnly: false = read-only, true otherwise 
        public WordprocessingDocument openWordDocument(string filePath, bool readOnly)
        {
            WordprocessingDocument document = WordprocessingDocument.Open(filePath, readOnly);
            return document;
        }

        // gets the body of an open .docx file
        // document: an instance of the .docx file opened using openWordDocument()
        public Body getWordDocumentBody(WordprocessingDocument document)
        {
            Body body = document.MainDocumentPart.Document.Body;
            return body;
        }

        public Table tableGetter(WordprocessingDocument document, int tableIndex)
        {
            Table table = document.MainDocumentPart.Document.Body.Elements<Table>().ElementAt(tableIndex);
            return table;
        }

        public TableRow tableRowGetter(Table table, int rowIndex)
        {
            TableRow row = table.Elements<TableRow>().ElementAt(rowIndex);
            return row;
        }

        public TableCell tableCellGetter(TableRow row, int cellIndex)
        {
            TableCell cell = row.Elements<TableCell>().ElementAt(cellIndex);
            return cell;
        }
        public TableCell DAIRCellGetter(WordprocessingDocument document, int tableIndex, int rowIndex, int cellIndex)
        {
            Table table = tableGetter(document, tableIndex);
            TableRow row = tableRowGetter(table, rowIndex);
            TableCell cell = tableCellGetter(row, cellIndex);
            return cell;
        }

        public string DAIRCellTextGetter(TableCell cell)
        {
            string text = cell.InnerText.Trim();
            if (text.Contains("FORMDROPDOWN"))
            {
                bool containsYes = text.IndexOf("yes", StringComparison.OrdinalIgnoreCase) >= 0;
                bool containsNo = text.IndexOf("no", StringComparison.OrdinalIgnoreCase) >= 0;
                if (containsYes)
                {
                    return "YES";
                }
                else if (containsNo)
                {
                    return "NO";
                }
                else if (text.Length > 12) {
                    string formText = "";
                    int index = text.IndexOf("FORMDROPDOWN");
                    if (index != -1)
                    {
                        formText = garbageCollector(text, "FORMDROPDOWN");
                    }
                    return formText;
                }
                else
                {
                    DropDownListFormField dropdown = cell.Elements<Paragraph>().First().Elements<Run>().First().Elements<FieldChar>().First().Elements<FormFieldData>().First().Elements<DropDownListFormField>().First();
                    var ddls = dropdown.DropDownListSelection;
                    if (ddls == null)
                    {
                        return "(empty)";
                    } else
                    {
                        int selectedIndex = dropdown.DropDownListSelection.Val;
                        ListEntryFormField selected = dropdown.Elements<ListEntryFormField>().ElementAt(selectedIndex);
                        string selectedText = selected.Val;
                        return selectedText;
                    }

                    
                }
            }

            else if (text.Contains("FORMTEXT"))
            {
                string cellText = "(empty)";
                if (text.Length > 8)
                {
                    int index = text.IndexOf("FORMTEXT");
                    if (index != -1)
                    {
                        cellText = garbageCollector(text, "FORMTEXT");
                    }
                }
                return cellText;

            }
            else
            {
                return cell.InnerText;
            }

            // gets rid of unwanted text in a string and returns the concatenated substrings
           
        }

        public void DAIRparser (string path)
        {
            DocumentProcessor d = new DocumentProcessor();
            int successCounter = 0;
            int errorCounter = 0;

            int year = DateTime.Now.Year;
            int month = DateTime.Now.Month;
            var now = DateTime.Now.ToString();
            now = now.Replace(":", " ");
            string currentPath = Directory.GetCurrentDirectory();
            string errorPath = currentPath + "\\Failed Documents";
            string logDirectory = currentPath + "\\Logs" + "\\" + year + "\\" + month;
            string logPath = logDirectory + "\\" + now + ".txt";


            foreach (string file in Directory.GetFiles(path))
            {
                string filePath = Path.GetFullPath(file);
                string fileName = Path.GetFileName(file);
                string errorDestination = Path.Combine(errorPath, fileName);
                string fileExtension = Path.GetExtension(file);
                if (fileExtension == ".docx")
                {
                    try
                    {
                        // there is code that tries to cover cases where OC does not use the dropdowns, otherwise the code would be much cleaner
                        WordprocessingDocument document = d.openWordDocument(filePath, true);
                        Body docBody = d.getWordDocumentBody(document);

                        Paragraph headerParagraph = docBody.Elements<Paragraph>().ElementAt(1);
                        string headerParagraphText = headerParagraph.InnerText.Trim();

                        Run airportNameRun = headerParagraph.Elements<Run>().ElementAt(3);
                        string airportName = airportNameRun.InnerText;

                        // if using dropdown
                        if (airportName == "")
                        {
                            int dateIndex = headerParagraphText.IndexOf("date", StringComparison.OrdinalIgnoreCase);
                            string tempAirportName = headerParagraphText.Substring(0, dateIndex);
                            bool usedDd = headerParagraphText.IndexOf("FORMDROPDOWN", StringComparison.OrdinalIgnoreCase) >= 0;
                            if (usedDd)
                            {
                                DropDownListFormField dropdown = airportNameRun.Elements<FieldChar>().First().Elements<FormFieldData>().First().Elements<DropDownListFormField>().First();
                                int selectedIndex = dropdown.DropDownListSelection.Val;
                                ListEntryFormField selected = dropdown.Elements<ListEntryFormField>().ElementAt(selectedIndex);
                                airportName = selected.Val;
                            }
                        }
                        Console.WriteLine("Airport: " + airportName);

                        SdtRun monthDayYear = headerParagraph.Elements<SdtRun>().ElementAt(0);
                        string monthDayYearText = monthDayYear.InnerText;
                        Console.WriteLine("Date: " + monthDayYearText);

                        string tempTimeText;
                        string timeText;
                        int timeIndex1 = headerParagraphText.IndexOf("time", StringComparison.OrdinalIgnoreCase);
                        int timeIndex2 = headerParagraphText.IndexOf("time:", StringComparison.OrdinalIgnoreCase);
                        bool withColon = timeIndex2 >= 0;
                        bool noColon = timeIndex1 >= 0;
                        int timeIndexEnd;
                        if (withColon)
                        {
                            timeIndexEnd = 5;
                            tempTimeText = headerParagraphText.Substring(timeIndex2);
                        }
                        else if (noColon)
                        {
                            timeIndexEnd = 4;
                            tempTimeText = headerParagraphText.Substring(timeIndex1);
                        }
                        else
                        {
                            timeIndexEnd = 4;
                            tempTimeText = "";
                        }

                        int garbageIndex = tempTimeText.IndexOf("FORMTEXT");
                        tempTimeText = tempTimeText.Trim();
                        if (garbageIndex >= 0)
                        {
                            timeText = (tempTimeText.Substring(timeIndexEnd, garbageIndex - timeIndexEnd) + tempTimeText.Substring(garbageIndex + 8)).Trim();
                        }
                        else
                        {
                            timeText = tempTimeText.Substring(timeIndexEnd).Trim();
                        }

                        Console.WriteLine("Time: " + timeText + "\n");





                        int numRows;
                        int numTables = docBody.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>().Count();
                        int numCells;
                        int tableCounter;
                        int rowCounter;
                        int cellCounter;

                        for (tableCounter = 0; tableCounter < numTables; tableCounter++)
                        {
                            numRows = docBody.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>().ElementAt(tableCounter).Elements<TableRow>().Count();
                            for (rowCounter = 0; rowCounter < numRows; rowCounter++)
                            {
                                numCells = docBody.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>().ElementAt(tableCounter).Elements<TableRow>().ElementAt(rowCounter).Elements<TableCell>().Count();
                                for (cellCounter = 0; cellCounter < numCells; cellCounter++)
                                {
                                    TableCell cell = d.DAIRCellGetter(document, tableCounter, rowCounter, cellCounter);
                                    string text = d.DAIRCellTextGetter(cell);
                                    Console.Write(text + " ");
                                }
                                Console.Write("\n");
                            }
                            Console.Write("\n");
                        }

                        string otherComments = "";
                        string version = "";
                        string completedBy = "";
                        // other comments
                        foreach (var text in docBody.Descendants<Text>())
                        {
                            if (text.Text.Contains("OTHER"))
                            {
                                Run run = (Run)text.Parent;
                                Paragraph para = (Paragraph)run.Parent;
                                otherComments = para.InnerText.Trim();
                                int garbage1 = otherComments.IndexOf(":");
                                int garbage2 = otherComments.IndexOf("OTHER COMMENTS", StringComparison.OrdinalIgnoreCase);
                                bool colon = (garbage1 != -1 && garbage1 <= 15);
                                bool otherExists = garbage2 != -1;
                                if (colon)
                                {
                                    otherComments = otherComments.Substring(garbage1 + 1).Trim();
                                }
                                else if (otherExists)
                                {
                                    otherComments = otherComments.Substring(garbage2 + 14).Trim();
                                }

                            }
                            if (text.Text.IndexOf("completed by", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                Run run = (Run)text.Parent;
                                Paragraph para = (Paragraph)run.Parent;
                                completedBy = para.InnerText.Trim();
                                //"Completed by:  ROBBIE NINGIURUVIKVersion: May 30, 2017"
                                int versionIndex = completedBy.IndexOf("version", StringComparison.OrdinalIgnoreCase);
                                bool versionExists = versionIndex >= 0;
                                if (versionExists)
                                {
                                    string temp = completedBy;
                                    string firstHalf = temp.Substring(0, versionIndex);
                                    int completedByIndex = firstHalf.IndexOf("Completed by", StringComparison.OrdinalIgnoreCase);
                                    completedBy = (firstHalf.Substring(0, completedByIndex) + firstHalf.Substring(completedByIndex)).Trim();
                                    completedBy = d.garbageCollector(completedBy, "Completed by");
                                    completedBy = d.garbageCollector(completedBy, ":");
                                    string secondHalf = temp.Substring(versionIndex);
                                    versionIndex = secondHalf.IndexOf("version", StringComparison.OrdinalIgnoreCase);
                                    version = secondHalf.Substring(0, versionIndex) + secondHalf.Substring(versionIndex);
                                    version = d.garbageCollector(version, "Version").Trim();
                                    version = d.garbageCollector(version, ":").Trim();
                                }
                                else
                                {
                                    completedBy = (d.garbageCollector(completedBy, "Completed by")).Trim();
                                    completedBy = (d.garbageCollector(completedBy, ":")).Trim();
                                    completedBy = (d.garbageCollector(completedBy, "FORMTEXT")).Trim();
                                }


                            }
                            else if (version == "" && text.Text.IndexOf("version", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                // TODO
                            }
                        }
                        otherComments = d.garbageCollector(otherComments, "FORMTEXT");
                        otherComments = d.garbageCollector(otherComments, "OTHER COMMENTS:").Trim();
                        Console.WriteLine("OTHER COMMENTS: " + otherComments);
                        Console.WriteLine("Completed By: " + completedBy);
                        Console.WriteLine("Version: " + version);
                        // commented out code for writing to word doc
                        //Paragraph para = docBody.AppendChild(new Paragraph());
                        //Run run = para.AppendChild(new Run());
                        //run.AppendChild(new Text("Append text in body, but text is not saved - OpenWordprocessingDocumentReadonly"));
                        //document.MainDocumentPart.Document.Save();
                        Console.WriteLine("-----------------------------------------------------");
                        successCounter++;
                    }
                    catch (FileNotFoundException e)
                    {
                        errorCounter++;
                        Logger.Logger logger = new Logger.Logger(logPath);
                        logger.Append(e.Message);
                        continue;

                    }
                    catch (IOException e)
                    {
                        errorCounter++;
                        Logger.Logger logger = new Logger.Logger(logPath);
                        logger.Append(e.Message);
                        File.Copy(file, errorDestination, true);
                        Environment.Exit(-1);
                    }
                    catch (System.Exception e)
                    {
                        errorCounter++;
                        Logger.Logger logger = new Logger.Logger(logPath);
                        logger.Append(e.Message);
                        File.Copy(file, errorDestination, true);
                        continue;
                    }
                }

            }
            MessageBox.Show("Success: " + successCounter + "        Failed: " + errorCounter + "\n Failed files are copied under /Failed Documents.", "Completion Report");
        }
        public string garbageCollector(string src, string garbage)
        {
            src.Trim();
            string clean;
            int garbageIndex = src.IndexOf(garbage);
            if (garbageIndex >= 0)
            {
                clean = src.Substring(0, garbageIndex) + src.Substring(garbageIndex + garbage.Length);
                return garbageCollector(clean.Trim(), garbage);
            } 
            else
            {
                return src;
            }
            
        }
    }
}
