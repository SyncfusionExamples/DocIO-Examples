using System;
using System.IO;
using System.Threading.Tasks;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;

namespace Word_to_PDF_Batch_process
{
    class Program
    {
        static void Main(string[] args)
        {
            //Gets multiple word document from the data folder
            var files = Directory.GetFiles(@"../../../Data/");
            Parallel.ForEach(files, file =>
            {
                //Open the file as Stream
                FileStream SourceStream = File.Open(file, FileMode.Open, FileAccess.Read);
                //Loads file stream into Word document
                WordDocument srcDoc = new WordDocument(SourceStream, FormatType.Automatic);
                //Instantiation of DocIORenderer for Word to PDF conversion
                DocIORenderer renderer = new DocIORenderer();
                //Converts Word document into PDF document
                PdfDocument pdfDocument = renderer.ConvertToPDF(srcDoc);
                //Saves the PDF file
                FileStream outputStream = new FileStream(@file + ".pdf", FileMode.OpenOrCreate, FileAccess.ReadWrite);
                pdfDocument.Save(outputStream);
                // Dispose the objects
                pdfDocument.Close(true);
                srcDoc.Close();
                renderer.Dispose();
                SourceStream.Dispose();
                outputStream.Dispose();
            });
            Console.WriteLine("Process Completed!!");
        }
    }
}
