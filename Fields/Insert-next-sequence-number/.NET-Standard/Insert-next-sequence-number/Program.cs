using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Insert_next_sequence_number
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Opens the template document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
                {
                    //Accesses sequence field in the document.
                    WParagraph paragraph = document.LastSection.Body.ChildEntities[2] as WParagraph;
                    WSeqField field = paragraph.ChildEntities[1] as WSeqField;
                    //Enables a flag to insert next number for sequence field.
                    field.InsertNextNumber = true;
                    //Updates the document fields.
                    document.UpdateDocumentFields();
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
}
