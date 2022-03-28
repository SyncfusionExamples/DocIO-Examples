using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;
using Syncfusion.DocToPDFConverter;
using Syncfusion.OfficeChartToImageConverter;
using Syncfusion.Pdf;
using System.IO;

namespace Adjust_image_quality
{
    class Program
    {
        static void Main(string[] args)
        {
            //Loads an existing Word document.
            using (WordDocument wordDocument = new WordDocument(Path.GetFullPath(@"../../Template.docx"), FormatType.Docx))
            {
                //Initializes the ChartToImageConverter for converting charts during Word to pdf conversion.
                wordDocument.ChartToImageConverter = new ChartToImageConverter();
                //Sets the scaling mode for charts (Normal mode reduces the Pdf file size).
                wordDocument.ChartToImageConverter.ScalingMode = ScalingMode.Normal;
                //Creates an instance of the DocToPDFConverter.
                using (DocToPDFConverter converter = new DocToPDFConverter())
                {
                    //Sets the image quality to reduce the Pdf file size.
                    converter.Settings.ImageQuality = 50;
                    //Converts Word document into PDF document.
                    using (PdfDocument pdfDocument = converter.ConvertToPDF(wordDocument))
                    {
                        //Saves the PDF file to file system.
                        pdfDocument.Save(Path.GetFullPath(@"../../WordToPDF.pdf"));
                    }
                }
            }
        }
    }
}
