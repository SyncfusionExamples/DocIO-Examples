using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;

namespace Open_and_read_Word_document
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the file as a stream.
            using (FileStream inputFileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Load the file stream into a Word document.
                using (WordDocument document = new WordDocument(inputFileStream, FormatType.Automatic))
                {
                    //Get the Word document text.
                    string text = document.GetText();
                    //Display Word document's text content.
                    Console.WriteLine(text);
                    Console.ReadLine();
                }
            }
        }
    }
}
