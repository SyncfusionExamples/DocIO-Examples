using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

namespace Check_document_has_diff_after_compare
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Load the original and revised Word documents.
            using (FileStream originalStream = new FileStream(Path.GetFullPath(@"Data/Original.docx"), FileMode.Open, FileAccess.Read))
            using (FileStream revisedStream = new FileStream(Path.GetFullPath(@"Data/Revised.docx"), FileMode.Open, FileAccess.Read))
            {
                using (WordDocument originalDocument = new WordDocument(originalStream, FormatType.Docx))
                using (WordDocument revisedDocument = new WordDocument(revisedStream, FormatType.Docx))
                {
                    // Configure comparison options to ignore formatting changes.
                    ComparisonOptions compareOptions = new ComparisonOptions();
                    compareOptions.DetectFormatChanges = false;

                    // Compare the documents.
                    originalDocument.Compare(revisedDocument);

                    // Check if there are content differences.
                    if (originalDocument.HasChanges)
                        Console.WriteLine("Differences detected in the document content.");
                    else
                        Console.WriteLine("The documents have the same content.");

                    Console.ReadLine();
                }
            }
        }
    }
}
