using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using Syncfusion.DocIORenderer;
using Syncfusion.Drawing.Fonts;

namespace Custom_font_registration
{
    internal class Program
    {
        static void Main(string[] args)
        {
            FontManager.RegisterFonts(Path.GetFullPath(@"Fonts"));
            List<string> registeredFontNames = FontManager.RegisteredFontNames;
            ConvertWordToImage();
            FontManager.ClearRegisteredFonts(true);
        }

        static void ConvertWordToImage()
        {
            // Load the Word document
            using (WordDocument wordDocument = new WordDocument(Path.GetFullPath(@"Data/Template.docx"), FormatType.Automatic))
            { 
                using (DocIORenderer renderer = new DocIORenderer())
                {
                    // Convert the entire Word document to images
                    Stream[] imageStreams = wordDocument.RenderAsImages();

                    for (int i = 0; i < imageStreams.Length; i++)
                    {
                        // Save each page as an image
                        string savePath = Path.GetFullPath($@"Output/Page_{i + 1}.jpeg");

                        using (FileStream outputStream = File.Create(savePath))
                        {
                            imageStreams[i].CopyTo(outputStream);
                        }
                    }
                }
             }   
        }
    }
}
