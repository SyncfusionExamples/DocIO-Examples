using System.Diagnostics;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

class Program
{
    static void Main()
    {
        using (FileStream sourceFileStream = new FileStream(Path.GetFullPath(@"Data/SourceDocument/Document-100.docx"), FileMode.Open, FileAccess.Read))
        {
            using (FileStream destinationFileStream = new FileStream(Path.GetFullPath(@"Data/DestinationDocument/Document-100.docx"), FileMode.Open, FileAccess.Read))
            {
                using (WordDocument mainDoc = new WordDocument(sourceFileStream, FormatType.Docx))
                {
                    using (WordDocument mergeDoc = new WordDocument(destinationFileStream, FormatType.Docx))
                    {
                        Stopwatch sw = Stopwatch.StartNew();
                        mainDoc.ImportContent(mergeDoc, ImportOptions.UseDestinationStyles);
                        sw.Stop();
                        Console.WriteLine("Time taken for Merge Documents:" + sw.Elapsed.TotalSeconds);
                        using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/MergedDocument.docx"), FileMode.Create))
                        {
                            mainDoc.Save(outputFileStream, FormatType.Docx);
                        }
                    }  
                }               
            }
        }
    }
}