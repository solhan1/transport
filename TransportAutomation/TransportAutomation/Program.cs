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
                        TableCell cell0 = d.DAIRCellGetter(document, 3, 0, 0);
                        TableCell cell1 = d.DAIRCellGetter(document, 3, 0, 1);
                        TableCell cell2 = d.DAIRCellGetter(document, 0, 2, 2);
                        string text0 = d.DAIRCellTextGetter(cell0);
                        string text1 = d.DAIRCellTextGetter(cell1);
                        string text2 = d.DAIRCellTextGetter(cell2);
                        //Table table = document.MainDocumentPart.Document.Body.Elements<Table>().ElementAt(3);
                        //TableRow row = table.Elements<TableRow>().ElementAt(0);
                        //TableCell cell0 = row.Elements<TableCell>().ElementAt(0);
                        //TableCell cell1 = row.Elements<TableCell>().ElementAt(1);
                        //DropDownListFormField dd = cell0.Elements<DropDownListFormField>();
                        Console.WriteLine(text0);
                        Console.WriteLine(text1);
                        Console.WriteLine(text2);
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
