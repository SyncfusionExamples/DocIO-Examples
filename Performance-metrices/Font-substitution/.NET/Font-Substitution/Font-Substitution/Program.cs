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
                    //Hooks the font substitution event
                    wordDocument.FontSettings.SubstituteFont += FontSettings_SubstituteFont;
                    //Creates an instance of DocIORenderer.
                    using (DocIORenderer renderer = new DocIORenderer())
                    {
                        //Unhooks the font substitution event after converting to PDF
                        wordDocument.FontSettings.SubstituteFont -= FontSettings_SubstituteFont;
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

    /// <summary>
    /// Handles font substitution when the original font used in the document
    /// is not available in the system.
    /// </summary>
    private static void FontSettings_SubstituteFont(object sender, SubstituteFontEventArgs args)
    {
        //Sets the alternate font when a specified font is not installed in the production environment
        //If "Arial Unicode MS" font is not installed, then it uses the "Arial" font
        //For other missing fonts, uses the "Times New Roman"
        if (args.OriginalFontName == "Arial Unicode MS")
            args.AlternateFontName = "Arial";
        else
            args.AlternateFontName = "Times New Roman";
    }
}
