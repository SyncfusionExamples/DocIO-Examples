using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChartToImageConverter;
using System.IO;
using Syncfusion.DocToPDFConverter;
using Syncfusion.Pdf;

namespace Convert_Word_document_to_PDF
{
    class Program
    {
        static void Main(string[] args)
        {
            //Load an existing Word document.
            using (WordDocument wordDocument = new WordDocument(Path.GetFullPath(@"../../Template.docx"), FormatType.Docx))
            {
                //Initialize the ChartToImageConverter for converting charts during Word to pdf conversion.
                wordDocument.ChartToImageConverter = new ChartToImageConverter();
                //Create an instance of DocToPDFConverter.
                using (DocToPDFConverter converter = new DocToPDFConverter())
                {
                    //Convert Word document into PDF document.
                    using (PdfDocument pdfDocument = converter.ConvertToPDF(wordDocument))
                    {
                        //Save the PDF file.
                        pdfDocument.Save(Path.GetFullPath(@"../../WordtoPDF.pdf"));
                    }
                }
            }
        }
    }
}
