using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Diagnostics;
using System.Reflection.Metadata;
using System.Xml.Linq;

namespace Reject_All_Changes_Word_document
{
    class Program
    {
        static void Main(string[] args)
        {
            //Load the document.
            using (FileStream inputFileStream = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open, FileAccess.Read))
            {
                //Load the word document.
                using (WordDocument wordDocument = new WordDocument(inputFileStream, FormatType.Docx))
                {
                    Stopwatch stopwatch = new Stopwatch();
                    stopwatch.Start();
                    //Reject all the tracked changes revisions in the Word document
                    if (wordDocument.HasChanges)
                        wordDocument.Revisions.RejectAll();
                    stopwatch.Stop();
                    Console.WriteLine($"Time taken for reject all changes in word Document: " + stopwatch.Elapsed.TotalSeconds);
                    //Create file stream
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to file stream
                        wordDocument.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
        