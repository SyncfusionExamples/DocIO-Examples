using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Add_editable_range
{
    class Program
    {
        static void Main(string[] args)
        {
            //Create a Word document
            using (WordDocument document = new WordDocument())
            {
                //Add a section and a paragraph to the Word document
                document.EnsureMinimal();
                WParagraph paragraph = document.LastParagraph;

                //Append text to the paragraph
                paragraph.AppendText("Adventure Works Cycles, the fictitious company on which the AdventureWorks ");

                //Add an editable range to the paragraph
                EditableRangeStart editableRangeStart = paragraph.AppendEditableRangeStart();
                paragraph.AppendText("sample databases are based, is a large, multinational manufacturing company.");
                paragraph.AppendEditableRangeEnd(editableRangeStart);

                //Set protection with a password to allow read-only access
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
