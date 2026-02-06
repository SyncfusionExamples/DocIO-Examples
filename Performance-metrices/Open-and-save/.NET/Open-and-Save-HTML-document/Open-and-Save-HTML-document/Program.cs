using System.Diagnostics;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;


class Program
{
    static void Main()
    {
        Stopwatch stopwatch = Stopwatch.StartNew();
        //Open a file as a stream.
        using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Input.html"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        {
            //Load the file stream into a HTML document.
            using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Html))
            {
                //Create a file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.html"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Save a HTML file to the file stream.
                    document.Save(outputFileStream, FormatType.Html);
                    stopwatch.Stop();
                    Console.WriteLine($"Time taken to open and save a 100-page HTML document: {stopwatch.Elapsed.TotalSeconds} seconds");
                }
            }
        }
    }
}
