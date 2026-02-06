using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Diagnostics;

namespace Compare_Word_documents
{
    class Program
    {
        static void Main(string[] args)
        {
            //Load the original document.
            using (FileStream originalDocumentStreamPath = new FileStream(Path.GetFullPath(@"Data/OriginalDocument.docx"), FileMode.Open, FileAccess.Read))
            {
                using (WordDocument originalDocument = new WordDocument(originalDocumentStreamPath, FormatType.Docx))
                {
                    //Load the revised document.
                    using (FileStream revisedDocumentStreamPath = new FileStream(Path.GetFullPath(@"Data/RevisedDocument.docx"), FileMode.Open, FileAccess.Read))
                    {
                        using (WordDocument revisedDocument = new WordDocument(revisedDocumentStreamPath, FormatType.Docx))
                        {
                            Stopwatch stopwatch = new Stopwatch();
                            stopwatch.Start();
                            // Compare the original and revised Word documents.
                            originalDocument.Compare(revisedDocument);
                            stopwatch.Stop();
                            Console.WriteLine($"Time taken for comapring Documents: " + stopwatch.Elapsed.TotalSeconds);
                            //Creates file stream.
                            using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                            {
                                //Saves the Word document to file stream.
                                originalDocument.Save(outputStream, FormatType.Docx);
                            }
                        }
                    }
                }
            }
        }
    }
}
        