using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using System.IO;

namespace Group_permission_for_editable_range
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
                WParagraph paragraph = document.LastParagraph;

                //Append text into the paragraph
                paragraph.AppendText("Adventure Works Cycles, the fictitious company on which the AdventureWorks ");

                //Adds an editable range to the paragraph
                EditableRangeStart editableRangeStart = paragraph.AppendEditableRangeStart();

                //Set the editor group
                editableRangeStart.EditorGroup = EditorType.Everyone;

                paragraph.AppendText("sample databases are based, is a large, multinational manufacturing company.");
                paragraph.AppendEditableRangeEnd(editableRangeStart);

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
