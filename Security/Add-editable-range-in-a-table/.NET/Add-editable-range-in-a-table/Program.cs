using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using System.IO;

namespace Add_editable_range_in_a_table
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load the Word document
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"Data/Template.docx")))
            {
                // Access the first table in the first section of the document
                WTable table = document.Sections[0].Tables[0] as WTable;
                // Access the paragraph in the third row and third column of the table
                WParagraph paragraph = table[2, 2].ChildEntities[0] as WParagraph;
                // Create a new editable range start for the table cell paragraph
                EditableRangeStart editableRangeStart = new EditableRangeStart(document);
                // Insert the editable range start at the beginning of the paragraph
                paragraph.ChildEntities.Insert(0, editableRangeStart);
                // Set the editor group for the editable range to allow everyone to edit
                editableRangeStart.EditorGroup = EditorType.Everyone;
                // Apply editable range to second column only
                editableRangeStart.FirstColumn = 1;
                editableRangeStart.LastColumn = 1;
                // Access the paragraph
                paragraph = table[5, 2].ChildEntities[0] as WParagraph;
                // Append an editable range end to close the editable region
                paragraph.AppendEditableRangeEnd();
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