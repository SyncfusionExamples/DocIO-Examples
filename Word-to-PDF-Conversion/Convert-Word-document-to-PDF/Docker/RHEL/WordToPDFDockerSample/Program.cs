using System;
using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;

namespace WordToPDFDockerSample
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the file as Stream
            FileStream docStream = new FileStream(@"Adventure.docx", FileMode.Open, FileAccess.Read);
            //Loads file stream into Word document
            WordDocument wordDocument = new WordDocument(docStream, Syncfusion.DocIO.FormatType.Automatic);
            docStream.Dispose();
            //Instantiation of DocIORenderer for Word to PDF conversion
            DocIORenderer render = new DocIORenderer();
            //Converts Word document into PDF document
            PdfDocument pdfDocument = render.ConvertToPDF(wordDocument);
            //Releases all resources used by the Word document and DocIO Renderer objects
            render.Dispose();
            wordDocument.Dispose();
            //Saves the PDF file
            FileStream outputStream = new FileStream("Output.pdf", FileMode.OpenOrCreate, FileAccess.ReadWrite);
            pdfDocument.Save(outputStream);
            //Closes the instance of PDF document object
            pdfDocument.Close();
            outputStream.Dispose();
        }
    }
}
