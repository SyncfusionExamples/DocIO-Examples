using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Office;
using Syncfusion.Pdf;
using Syncfusion.Office;
namespace Modify_the_exiting_fallback_fonts
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Opens the file as stream.
            using FileStream inputStream = new FileStream(@"../../../Template.docx", FileMode.Open, FileAccess.Read);
            //Loads an existing Word document file stream.
            using WordDocument wordDocument = new WordDocument(inputStream, Syncfusion.DocIO.FormatType.Docx);
            //Initialize the default fallback fonts collection.
            wordDocument.FontSettings.FallbackFonts.InitializeDefault();
            FallbackFonts fallbackFonts = wordDocument.FontSettings.FallbackFonts;
            foreach (FallbackFont fallbackFont in fallbackFonts)
            {
                //Customize a default fallback font name as "David" for the Hebrew script.
                if (fallbackFont.ScriptType == ScriptType.Hebrew)
                    fallbackFont.FontNames = "David";
                //Customize a default fallback font name as "Microsoft Sans Serif" for the Thai script.
                else if (fallbackFont.ScriptType == ScriptType.Thai)
                    fallbackFont.FontNames = "Microsoft Sans Serif";
            }
            //Instantiation of DocIORenderer for Word to PDF conversion.
            using DocIORenderer render = new DocIORenderer();
            //Converts Word document into PDF document.
            using PdfDocument pdfDocument = render.ConvertToPDF(wordDocument);
            //Saves the PDF file to file system.
            using FileStream outputStream = new FileStream(@"../../../WordToPDF.pdf", FileMode.OpenOrCreate, FileAccess.ReadWrite);
            pdfDocument.Save(outputStream);
        }
    }
}
