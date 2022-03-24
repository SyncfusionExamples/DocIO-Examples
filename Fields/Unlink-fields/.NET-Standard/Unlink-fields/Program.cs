using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Unlink_fields
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates an instance of WordDocument class.
            using (WordDocument document = new WordDocument())
            {
                //Adds a new section into the Word Document.
                IWSection section = document.AddSection();
                //Adds a new paragraph into Word document and appends text into paragraph.
                IWParagraph paragraph = section.AddParagraph();
                paragraph.AppendText("Today's Date: ");
                //Adds the new Date field in Word document with field name and its type.
                WField field = paragraph.AppendField("Date", FieldType.FieldDate) as WField;
                //Updates the field.
                field.Update();
                //Unlink the field.
                field.Unlink();
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
