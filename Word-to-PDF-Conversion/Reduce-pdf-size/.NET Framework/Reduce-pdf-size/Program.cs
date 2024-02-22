using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocToPDFConverter;
using Syncfusion.Pdf;

namespace Reduce_pdf_size
{
    internal class Program
    {
        static void Main(string[] args)
        {
            WordDocument wordDocument = new WordDocument(@"../../Data/Template.docx", FormatType.Docx);
            DocToPDFConverter converter = new DocToPDFConverter();

            // Adjust image quality and resolution
            converter.Settings.ImageQuality = 50;
            converter.Settings.ImageResolution = 640;

            // Convert Word document to PDF
            PdfDocument pdfDocument = converter.ConvertToPDF(wordDocument);

            // Set PDF compression level
            pdfDocument.Compression = PdfCompressionLevel.Best;

            // Save the PDF document
            pdfDocument.Save("../../Data/Output.pdf");

            // Close document instances
            wordDocument.Close();
            pdfDocument.Close();
        }
    }
}
