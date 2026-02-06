using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Diagnostics;

namespace Find_and_replace_in_a_worddocument
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Open the template Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    Stopwatch stopwatch = new Stopwatch();
                    stopwatch.Start();
                    //Find all occurrences of a misspelled word and replaces with properly spelled word
                    int replacedCount = document.Replace("document", "DocIO", false, false);
                    stopwatch.Stop();
                    Console.WriteLine(replacedCount);
                    Console.WriteLine($"Time taken for Replace (string):" + stopwatch.Elapsed.TotalSeconds);
                    //Create file stream
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to file stream
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }       
    }
}
