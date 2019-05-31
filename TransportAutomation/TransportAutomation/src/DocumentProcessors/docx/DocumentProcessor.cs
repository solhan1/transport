using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.IO.Packaging;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

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
                    int selectedIndex = dropdown.DropDownListSelection.Val;
                    ListEntryFormField selected = dropdown.Elements<ListEntryFormField>().ElementAt(selectedIndex);
                    string selectedText = selected.Val;
                    return selectedText;
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
