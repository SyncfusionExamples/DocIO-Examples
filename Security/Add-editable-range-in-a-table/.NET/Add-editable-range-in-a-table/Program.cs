using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using System.IO;

namespace Add_editable_range_in_a_table
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a Word document
            using (WordDocument document = new WordDocument())
            {
                //Adds a section and a paragraph in the Word document
                document.EnsureMinimal();

                //Adds a table
                WTable table = document.LastSection.AddTable() as WTable;
                table.ResetCells(2, 3);

                //Access each table cell and append text
                table[0, 0].AddParagraph().AppendText("Row1 Col1");
                table[0, 1].AddParagraph().AppendText("Row1 Col2");
                table[0, 2].AddParagraph().AppendText("Row1 Col3");
                table[1, 0].AddParagraph().AppendText("Row2 Col1");
                table[1, 1].AddParagraph().AppendText("Row2 Col2");
                table[1, 2].AddParagraph().AppendText("Row2 Col3");

                //Starts the editable range in a table cell
                EditableRangeStart editableRangeStart = table[0, 1].Paragraphs[0].AppendEditableRangeStart();

                //Sets the first column where the editable range starts within a table
                editableRangeStart.EditableRange.FirstColumn = 1;

                //Ends the ediatble range in a table cell
                EditableRangeEnd rangeEnd = table[1, 2].Paragraphs[0].AppendEditableRangeEnd(editableRangeStart);

                //Sets the last column where the editable range ends within a table
                editableRangeStart.EditableRange.LastColumn = 2;

                //Sets the protection with password and allows only reading
                document.Protect(ProtectionType.AllowOnlyReading, "password");

                //Creates file stream
                using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream
                    document.Save(outputStream, FormatType.Docx);
                }
            }
        }
    }
}