using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using System;
using System.Diagnostics;

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
            string documentOutputFolder = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(fileName));
            Directory.CreateDirectory(documentOutputFolder);

            Stopwatch stopwatch = Stopwatch.StartNew();

            try
            {
                // Open an existing Word document
                WordDocument document = new WordDocument(inputPath);
                DocIORenderer renderer = new DocIORenderer();
                //Convert the entire Word document to images.
                Stream[] imageStreams = document.RenderAsImages();
                for (int i = 0; i < imageStreams.Length; i++)
                {
                    //Save the image stream as file.
                    string imagePath = Path.Combine(documentOutputFolder, $"WordToImage_{i + 1}.jpg");
                    using (FileStream fileStreamOutput = new(imagePath, FileMode.Create, FileAccess.Write))
                    {
                        imageStreams[i].CopyTo(fileStreamOutput);
                    }
                }
                stopwatch.Stop();
                Console.WriteLine($"{fileName} processed in {stopwatch.Elapsed.TotalSeconds} seconds");
                
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing {fileName}: {ex.Message}");
            }
        }
    }
}

