using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Open_and_save_Word_document
{
    class Program
    {
        static void Main(string[] args)
        {            
            using (FileStream inputFileStream = new FileStream(Path.GetFullPath(@"../../../HelloWorld.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an existing document from stream through constructor of `WordDocument` class.
                using (WordDocument document = new WordDocument(inputFileStream, FormatType.Automatic))
                {
                    //Appends text to the last paragraph of the document.
                    document.LastParagraph.AppendText("Hello World");
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                    //Closes the document.
                    document.Close();
                }
            }
        }
    }
}
