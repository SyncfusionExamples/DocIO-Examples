using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Add_field_in_Word_document
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates an instance of WordDocument class (Empty Word Document).
            using (WordDocument document = new WordDocument())
            {
                //Adds a new section to the Word Document.
                IWSection section = document.AddSection();
                //Adds a new paragraph to Word document and appends text into paragraph.
                IWParagraph paragraph = section.AddParagraph();
                paragraph.AppendText("Today's Date: ");
                //Adds the new Date field to Word document with field name and its type.
                WField field = paragraph.AppendField("Date", FieldType.FieldDate) as WField;
                //Field code used to describe how to display the date.
                field.FieldCode = @"DATE  \@" + "\"MMMM d, yyyy\"";
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
