
using System.Diagnostics;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;


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
            string outputPath = Path.Combine(outputFolder, fileName);

            Stopwatch stopwatch = Stopwatch.StartNew();

            try
            {
                // Load or open Word document
                WordDocument document = new WordDocument(inputPath);

                // Save the Word document to Output folder
                document.Save(outputPath);
                stopwatch.Stop();
                Console.WriteLine($"{fileName} open and saved in {stopwatch.Elapsed.TotalSeconds} seconds");
                document.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing {fileName}: {ex.Message}");
            }
        }
    }
}
