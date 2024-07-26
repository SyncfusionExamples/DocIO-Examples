using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Create_merge_field
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates an instance of a WordDocument.
            using (WordDocument document = new WordDocument())
            {
                //Adds a section and a paragraph in the document.
                document.EnsureMinimal();
                //Appends merge field to the last paragraph.
                document.LastParagraph.AppendField("FullName", FieldType.FieldMergeField);
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
