using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Remove_encryption_from_Word_document
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Opens an encrypted Word document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Docx, "syncfusion"))
                {
                    //Removes encryption in Word document.
                    document.RemoveEncryption();
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
