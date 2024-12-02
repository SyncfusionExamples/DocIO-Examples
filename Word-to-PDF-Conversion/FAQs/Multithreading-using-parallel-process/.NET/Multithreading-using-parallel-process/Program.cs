using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;
using System.Threading.Tasks;

namespace Multithreading_using_parallel_process
{
    class MultiThreading
    {
        static void Main(string[] args)
        {
            //Indicates the number of threads to be create.
            int limit = 5;
            Console.WriteLine("Parallel For Loop");
            Parallel.For(0, limit, count =>
            {
                Console.WriteLine("Task {0} started", count);
                //Create multiple Word document, one document on each thread.
                ConvertWordToPDF(count);
                Console.WriteLine("Task {0} is done", count);
            });
        }
        //Convert a Word document to PDF using multi-threading.
        static void ConvertWordToPDF(int count)
        {
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.docx"), FileMode.Open, FileAccess.Read))
            {
                // Load the Word document.
                using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
                {
                    // Convert Word document to PDF.
                    using (DocIORenderer renderer = new DocIORenderer())
                    {
                        PdfDocument pdfDocument = renderer.ConvertToPDF(document);
                        // Save the PDF document.
                        using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output" + count + ".pdf"), FileMode.Create, FileAccess.Write))
                        {
                            pdfDocument.Save(outputFileStream);
                        }
                        // Dispose pdf to free resources.
                        pdfDocument.Dispose();
                    }
                }
            }
        }
    }
}
