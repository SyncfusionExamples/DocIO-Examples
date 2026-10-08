using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using Syncfusion.Drawing.Fonts;

namespace Custom_font_registration
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            FontManager.RegisterFonts(Path.GetFullPath(@"Fonts"));
            List<string> registeredFontNames = FontManager.RegisteredFontNames;
            ConvertWordToPDF();
            FontManager.ClearRegisteredFonts(true);
        }

        private static void ConvertWordToPDF()
        {
            // Loads an existing Word document
            using (WordDocument wordDocument = new WordDocument(Path.GetFullPath(@"Data/Template.docx"), FormatType.Automatic))
            {
                // Creates an instance of DocIORenderer
                using (DocIORenderer renderer = new DocIORenderer())
                {
                    // Converts Word document into PDF document
                    using (PdfDocument pdfDocument = renderer.ConvertToPDF(wordDocument))
                    {
                        // Saves the PDF document
                        pdfDocument.Save(Path.GetFullPath(@"Output/Output.pdf"));
                    }
                }
            }
        }
    }
}
