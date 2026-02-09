using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using System.Diagnostics;
using System.IO;
using System.Reflection.Metadata;

class Program
{
    static void Main()
    {
        Stopwatch stopwatch = Stopwatch.StartNew();
        try
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                // Open and convert Word to PDF
                using (WordDocument document = new WordDocument(fileStreamPath,FormatType.Docx))
                {
                    DocIORenderer render = new DocIORenderer();
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.pdf"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        // Convert Word to PDF
                        using (PdfDocument pdfDocument = render.ConvertToPDF(document))
                        {
                            pdfDocument.Save(outputFileStream);
                        }
                        stopwatch.Stop();
                        Console.WriteLine($"Time taken to convert as PDF: {stopwatch.Elapsed.TotalSeconds} seconds");
                    }
                }
            }            
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Time taken to convert as PDF: {ex.Message}");
        }    
    }
}


