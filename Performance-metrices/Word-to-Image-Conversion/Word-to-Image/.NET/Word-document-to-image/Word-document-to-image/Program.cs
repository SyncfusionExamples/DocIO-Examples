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
                WordDocument document = new WordDocument(fileStreamPath,FormatType.Docx);
                //Initialize the default fallback fonts collection.
                document.FontSettings.FallbackFonts.InitializeDefault();
                DocIORenderer renderer = new DocIORenderer();
                //Convert the entire Word document to images.
                Stream[] imageStreams = document.RenderAsImages();
                for (int i = 0; i < imageStreams.Length; i++)
                {
                    //Save the image stream as file.
                    string imagePath = Path.Combine(Path.GetFullPath(@"Output/Images"), $"WordToImage_{i + 1}.jpg");
                    using (FileStream fileStreamOutput = new(imagePath, FileMode.Create, FileAccess.Write))
                    {
                        imageStreams[i].CopyTo(fileStreamOutput);
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
}

