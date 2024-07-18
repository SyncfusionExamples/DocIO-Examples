using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Open_encrypted_Word_document
{
    class Program
    {
        static void Main(string[] args)
        {
            //Opens an existing document from stream.
            using (FileStream inputFileStream = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an encrypted Word document.
                using (WordDocument document = new WordDocument(inputFileStream, "syncfusion"))
                {
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
}
