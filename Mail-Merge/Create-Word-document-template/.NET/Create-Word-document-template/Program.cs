using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Create_Word_document_template
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates an instance of a WordDocument.
            using (WordDocument document = new WordDocument())
            {
                //Adds one section and one paragraph to the document.
                document.EnsureMinimal();
                //Sets page margins to the last section of the document.
                document.LastSection.PageSetup.Margins.All = 72;
                //Appends text to the last paragraph.
                document.LastParagraph.AppendText("EmployeeId: ");
                //Appends merge field to the last paragraph.
                document.LastParagraph.AppendField("EmployeeId", FieldType.FieldMergeField);
                document.LastParagraph.AppendText("\nName: ");
                document.LastParagraph.AppendField("Name", FieldType.FieldMergeField);
                document.LastParagraph.AppendText("\nPhone: ");
                document.LastParagraph.AppendField("Phone", FieldType.FieldMergeField);
                document.LastParagraph.AppendText("\nCity: ");
                document.LastParagraph.AppendField("City", FieldType.FieldMergeField);
                //Creates file stream.
                using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputStream, FormatType.Docx);
                }
            }
        }
    }
}
