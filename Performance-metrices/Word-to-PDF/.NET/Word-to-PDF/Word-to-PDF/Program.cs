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
        string inputFolder = Path.GetFullPath("../../../Data/");
        string outputFolder = Path.GetFullPath("../../../Output/");

        Directory.CreateDirectory(outputFolder);

        // Get all .docx files in the Data folder
        string[] files = Directory.GetFiles(inputFolder, "*.docx");

        foreach (string inputPath in files)
        {
            string fileName = Path.GetFileName(inputPath);
            string outputPath = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(fileName) + ".pdf");

            Stopwatch stopwatch = Stopwatch.StartNew();

            try
            {
                // Open and convert Word to PDF
                using (WordDocument document = new WordDocument(inputPath))
                {
                    DocIORenderer render = new DocIORenderer();
                    // Convert Word to PDF
                    using (PdfDocument pdfDocument = render.ConvertToPDF(document))
                    {
                        pdfDocument.Save(outputPath);
                    }                    
                }
                stopwatch.Stop();
                Console.WriteLine($"{fileName} taken time to convert as PDF: {stopwatch.Elapsed.TotalSeconds} seconds");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing {fileName}: {ex.Message}");
            }
        }
    }
}


