using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
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
                // Open an existing Word document
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    //Hooks the font substitution event
                    document.FontSettings.SubstituteFont += FontSettings_SubstituteFont;
                    using (DocIORenderer renderer = new DocIORenderer())
                    {
                        //Convert the entire Word document to images.
                        Stream[] imageStreams = document.RenderAsImages();
                        //Unhooks the font substitution event after converting to PDF
                        document.FontSettings.SubstituteFont -= FontSettings_SubstituteFont;
                        for (int i = 0; i < imageStreams.Length; i++)
                        {
                            //Save the image stream as file.
                            string imagePath = Path.Combine(Path.GetFullPath(@"Output/Images"), $"WordToImage_{i + 1}.jpg");
                            using (FileStream fileStreamOutput = new(imagePath, FileMode.Create, FileAccess.Write))
                            {
                                imageStreams[i].CopyTo(fileStreamOutput);
                            }
                        }
                    }               
                }                  
                stopwatch.Stop();
                Console.WriteLine($"Word To Image processed in {stopwatch.Elapsed.TotalSeconds} seconds");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing Word to Image: {ex.Message}");
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


