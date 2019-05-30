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
            if (cell.InnerText.Contains("FORMDROPDOWN"))
            {
                DropDownListFormField dropdown = cell.Elements<Paragraph>().First().Elements<Run>().First().Elements<FieldChar>().First().Elements<FormFieldData>().First().Elements<DropDownListFormField>().First();
                int selectedIndex = dropdown.DropDownListSelection.Val;
                ListEntryFormField selected = dropdown.Elements<ListEntryFormField>().ElementAt(selectedIndex);
                string selectedText = selected.Val;
                return selectedText;
            }

            else if (cell.InnerText.Contains("FORMTEXT"))
            {
                return "(empty)";
            }
            else
            {
                return cell.InnerText;
            }
        }
    }
}
