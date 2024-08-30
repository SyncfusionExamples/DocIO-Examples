using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Refer_bookmark_in_SEQ_field
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Opens the template document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
                {
                    //Accesses sequence field in the document.
                    WParagraph paragraph = document.LastSection.Body.ChildEntities[5] as WParagraph;
                    WSeqField seqField = paragraph.ChildEntities[5] as WSeqField;
                    //Adds bookmark reference to the sequence field.
                    seqField.BookmarkName = "BkmkPurchase";
                    //Accesses sequence field in the document.
                    paragraph = document.LastSection.Body.ChildEntities[6] as WParagraph;
                    seqField = paragraph.ChildEntities[1] as WSeqField;
                    //Adds bookmark reference to the sequence field.
                    seqField.BookmarkName = "BkmkUnitPrice";
                    //Updates the document fields.s
                    document.UpdateDocumentFields();
                    //Creates file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
