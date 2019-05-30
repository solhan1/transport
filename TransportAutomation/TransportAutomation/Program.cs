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
                        WordprocessingDocument document = d.openWordDocument(fileName, true);
                        // Body docBody = d.getWordDocumentBody(document);
                        int tableCounter;
                        int rowCounter;
                        int numRows;
                        int numTables = 5;
                        int cellCounter;
                        int numCells = 4;

                        for (tableCounter = 0; tableCounter < numTables; tableCounter++)
                        {
                            if (tableCounter == 0 || tableCounter == 1 || tableCounter == 4)
                            {
                                numRows = 8;
                            }
                            else if (tableCounter == 2)
                            {
                                numRows = 6;
                            } 
                            else
                            {
                                numRows = 5;
                            }
                            for (rowCounter = 0; rowCounter < numRows; rowCounter++)
                            {
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
                        
                        // commented out code for writing to word doc
                        //Paragraph para = docBody.AppendChild(new Paragraph());
                        //Run run = para.AppendChild(new Run());
                        //run.AppendChild(new Text("Append text in body, but text is not saved - OpenWordprocessingDocumentReadonly"));
                        //document.MainDocumentPart.Document.Save();
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
