
using System.Diagnostics;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;


class Program
{
    static void Main()
    {
        Stopwatch stopwatch = Stopwatch.StartNew();
        //Open a file as a stream.
        using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Input.rtf"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        {
            //Load the file stream into a RTF document.
            using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Rtf))
            {
                //Create a file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.rtf"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Save a RTF file to the file stream.
                    document.Save(outputFileStream, FormatType.Rtf);
                    stopwatch.Stop();
                    Console.WriteLine($"Time taken to open and save a 100-page RTF document: {stopwatch.Elapsed.TotalSeconds} seconds");
                }
            }
        }
    }
}
