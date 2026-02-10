using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using System.Diagnostics;

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
                using (WordDocument wordDocument = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    //Creates an instance of DocIORenderer.
                    using (DocIORenderer renderer = new DocIORenderer())
                    {
                        //Sets ExportBookmarks for preserving Word document headings as PDF bookmarks
                        renderer.Settings.ExportBookmarks = Syncfusion.DocIO.ExportBookmarkType.Headings;
                        //Converts Word document into PDF document.
                        using (PdfDocument pdfDocument = renderer.ConvertToPDF(wordDocument))
                        {
                            using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.pdf"), FileMode.Create, FileAccess.ReadWrite))
                            {
                                pdfDocument.Save(outputFileStream);
                            }
                        }
                    }
                }
                stopwatch.Stop();
                Console.WriteLine($"Input.docx taken time to convert as PDF: {stopwatch.Elapsed.TotalSeconds} seconds");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing Input.docx: {ex.Message}");
        }
    }
}
