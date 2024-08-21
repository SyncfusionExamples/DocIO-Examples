using System.IO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Drawing;
using Syncfusion.Pdf;

namespace Use_alternate_font_without_installing
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open))
            {
                //Loads an existing Word document.
                using (WordDocument wordDocument = new WordDocument(fileStream, Syncfusion.DocIO.FormatType.Automatic))
                {
                    //Hooks the font substitution event.
                    wordDocument.FontSettings.SubstituteFont += FontSettings_SubstituteFont;
                    //Creates an instance of DocIORenderer.
                    using (DocIORenderer renderer = new DocIORenderer())
                    {
                        //Converts Word document into PDF document.
                        using (PdfDocument pdfDocument = renderer.ConvertToPDF(wordDocument))
                        {
                            //Unhooks the font substitution event after converting to PDF.
                            wordDocument.FontSettings.SubstituteFont -= FontSettings_SubstituteFont;
                            //Saves the PDF file to file system.    
                            using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Output.pdf"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
                            {
                                pdfDocument.Save(outputStream);
                            }
                        }
                    }
                }
            }
        }
        private static void FontSettings_SubstituteFont(object sender, SubstituteFontEventArgs args)
        {
            //Sets the alternate font when a specified font is not installed in the production environment.
            if (args.OriginalFontName == "Arial Unicode MS" && args.FontStyle == FontStyle.Regular)
                args.AlternateFontStream = new FileStream(Path.GetFullPath(@"Data/Arial.TTF"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            else
                args.AlternateFontName = "Times New Roman";
        }
    }
}
