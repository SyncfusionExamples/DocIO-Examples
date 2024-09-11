using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;

namespace Initialize_default_fallback_fonts
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Opens the file as stream.
            using FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read);
            //Loads an existing Word document file stream.
            using WordDocument wordDocument = new WordDocument(inputStream, Syncfusion.DocIO.FormatType.Docx);
            //Initialize the default fallback fonts collection.
            wordDocument.FontSettings.FallbackFonts.InitializeDefault();
            //Instantiation of DocIORenderer for Word to PDF conversion.
            using DocIORenderer render = new DocIORenderer();
            //Converts Word document into PDF document.
            using PdfDocument pdfDocument = render.ConvertToPDF(wordDocument);
            //Saves the PDF file to file system.
            using FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Output.pdf"), FileMode.OpenOrCreate, FileAccess.ReadWrite);
            pdfDocument.Save(outputStream);
        }
    }
}
