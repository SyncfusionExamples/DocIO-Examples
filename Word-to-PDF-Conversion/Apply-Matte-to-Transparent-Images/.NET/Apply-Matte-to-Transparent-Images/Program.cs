using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;

namespace Apply_mattee_to_transparent_images
{
    class Program
    {
        public static void Main(string[] args)
        {
            FileStream fileStream = new FileStream(Path.GetFullPath(@"Data\Template.docx"), FileMode.Open);
            //Loads an existing Word document
            WordDocument wordDocument = new WordDocument(fileStream, FormatType.Docx);
            //Instantiates DocIORenderer instance for Word to PDF conversion
            DocIORenderer renderer = new DocIORenderer();
            //Set to true to apply a matte color to transparent images.
            renderer.Settings.ApplyMatteToTransparentImages = true;
            //Converts Word document into PDF document
            PdfDocument pdfDocument = renderer.ConvertToPDF(wordDocument);
            //Closes the instance of Word document object
            wordDocument.Close();
            //Releases the resources occupied by DocIORenderer instance
            renderer.Dispose();
            //Saves the PDF file  
            pdfDocument.Save(Path.GetFullPath(@"../../../Output/Result.pdf"));
            //Closes the instance of PDF document object
            pdfDocument.Close();
        }
    }
}