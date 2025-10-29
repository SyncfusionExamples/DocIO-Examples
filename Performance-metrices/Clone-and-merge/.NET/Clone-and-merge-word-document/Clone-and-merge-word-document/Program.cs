using System.ComponentModel;
using System.Diagnostics;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

class Program
{
    static void Main()
    {
        string sourceFolder = Path.GetFullPath("../../../Data/SourceDocument/");
        string destinationFolder = Path.GetFullPath("../../../Data/DestinationDocument/");
        string outputFolder = Path.GetFullPath("../../../Output/");

        Directory.CreateDirectory(outputFolder);

        // Get all source files
        string[] sourceFiles = Directory.GetFiles(sourceFolder, "*.docx");

        foreach (string sourcePath in sourceFiles)
        {
            string fileName = Path.GetFileName(sourcePath);
            string destinationPath = Path.Combine(destinationFolder, fileName);

            if (!File.Exists(destinationPath))
            {
                Console.WriteLine($"Skipping {fileName} - No matching destination file found.");
                continue;
            }

            string outputPath = Path.Combine(outputFolder, fileName);

            Stopwatch stopwatch = Stopwatch.StartNew();

            try
            {
                // Open source and destination document
                WordDocument sourceDocument = new WordDocument(sourcePath); 
                WordDocument destinationDocument =  new WordDocument(destinationPath);

                // Clone and merge all slides
                foreach (WSection section in sourceDocument.Sections)
                {
                    WSection clonedSection = section.Clone();
                    destinationDocument.Sections.Add(clonedSection);
                }

                // Save the merged document.
                destinationDocument.Save(outputPath);

                stopwatch.Stop();
                Console.WriteLine($"{fileName} is cloned and merged in {stopwatch.Elapsed.TotalSeconds} seconds.");
                sourceDocument.Close();
                destinationDocument.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing {fileName}: {ex.Message}");
            }
        }
    }
}
