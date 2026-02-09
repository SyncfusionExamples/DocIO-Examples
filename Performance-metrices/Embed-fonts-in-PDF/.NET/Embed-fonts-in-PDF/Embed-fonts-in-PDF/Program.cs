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
                using (WordDocument wordDocument = new WordDocument(fileStreamPath,FormatType.Docx))
                {
                    //Creates an instance of DocIORenderer.
                    using (DocIORenderer renderer = new DocIORenderer())
                    {
                        //Sets true to embed complete TrueType fonts
                        renderer.Settings.EmbedCompleteFonts = true;
                        using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.pdf"), FileMode.Create, FileAccess.ReadWrite))
                        {
                            //Converts Word document into PDF document.
                            using (PdfDocument pdfDocument = renderer.ConvertToPDF(wordDocument))
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
