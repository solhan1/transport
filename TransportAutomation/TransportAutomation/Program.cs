using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TransportAutomation.src.DocumentProcessors.docx;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Office.CustomUI;
using System.IO;
using System.Windows.Forms;


namespace TransportAutomation
{
    class Program
    {
        static void Main(string[] args)
        {
            DocumentProcessor d = new DocumentProcessor();
            string path = Directory.GetCurrentDirectory();
            foreach (string fileName in Directory.GetFiles(path))
            {
                string fileExtension = Path.GetExtension(fileName);
                if (fileExtension == ".docx")
                {
                    try
                    {
                        // there is code that tries to cover cases where OC does not use the dropdowns, otherwise the code would be much cleaner
                        WordprocessingDocument document = d.openWordDocument(fileName, true);
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
                        } else if (noColon) {
                            timeIndexEnd = 4;
                            tempTimeText = headerParagraphText.Substring(timeIndex1);
                        } else
                        {
                            timeIndexEnd = 4;
                            tempTimeText = "";
                        }
                        
                        int garbageIndex = tempTimeText.IndexOf("FORMTEXT");
                        tempTimeText = tempTimeText.Trim();
                        if (garbageIndex >= 0)
                        {
                            timeText = (tempTimeText.Substring(timeIndexEnd, garbageIndex-timeIndexEnd) + tempTimeText.Substring(garbageIndex + 8)).Trim();
                        }
                        else
                        {
                            timeText = tempTimeText.Substring(timeIndexEnd).Trim();
                        }
                            
                        Console.WriteLine("Time: " + timeText + "\n");
                        



                        
                        int numRows;
                        int numTables = docBody.Elements<Table>().Count();
                        int numCells;
                        int tableCounter;
                        int rowCounter;
                        int cellCounter;

                        for (tableCounter = 0; tableCounter < numTables; tableCounter++)
                        {
                            numRows = docBody.Elements<Table>().ElementAt(tableCounter).Elements<TableRow>().Count();
                            for (rowCounter = 0; rowCounter < numRows; rowCounter++)
                            {
                                numCells = docBody.Elements<Table>().ElementAt(tableCounter).Elements<TableRow>().ElementAt(rowCounter).Elements<TableCell>().Count();
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
                                    otherComments = otherComments.Substring(garbage1+1).Trim();
                                }
                                else if (otherExists) {
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
                                } else
                                {
                                    completedBy = (d.garbageCollector(completedBy, "Completed by")).Trim();
                                    completedBy = (d.garbageCollector(completedBy, ":")).Trim();
                                    completedBy = (d.garbageCollector(completedBy, "FORMTEXT")).Trim();
                                }

                            
                            } else if (version == "" && text.Text.IndexOf("version", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                // TODO
                            }
                        }
                        otherComments = d.garbageCollector(otherComments, "FORMTEXT");
                        otherComments = d.garbageCollector(otherComments, "OTHER COMMENTS:").Trim();
                        Console.WriteLine("OTHER COMMENTS: " + otherComments);
                        Console.WriteLine("Completed By: " + completedBy);
                        Console.WriteLine("Version: " + version);
                        /*Paragraph completedByParagraph = docBody.Elements<Paragraph>().ElementAt(12);
                        string completedByParagraphText = completedByParagraph.InnerText;
                        int index = completedByParagraphText.IndexOf("Version");
                        string completedByText;
                        string versionText;
                        if (index != -1)
                        {
                            completedByText = completedByParagraphText.Substring(0, index);
                            versionText = completedByParagraphText.Substring(index);
                            Console.WriteLine(completedByText);
                            Console.WriteLine(versionText);
                        } */
                        // commented out code for writing to word doc
                        //Paragraph para = docBody.AppendChild(new Paragraph());
                        //Run run = para.AppendChild(new Run());
                        //run.AppendChild(new Text("Append text in body, but text is not saved - OpenWordprocessingDocumentReadonly"));
                        //document.MainDocumentPart.Document.Save();
                        Console.WriteLine("-----------------------------------------------------");
                    }
                    catch (FileNotFoundException e)
                    {
                        MessageBox.Show("ERROR: The file:" + fileName + "was not found");
                    }
                    catch (IOException e)
                    {
                        MessageBox.Show("ERROR: Documents to be processed must not be open. Please close them and try again.");
                        Environment.Exit(-1);
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show("ERROR: Uncaught error, please contact IT. " + e);
                    }
                }
                
            }
            
            
        }
    }
}
