using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Office;

namespace Fallback_fonts_for_custom_range
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Opens the file as stream
            using FileStream inputStream = new FileStream(@"../../../Template.docx", FileMode.Open, FileAccess.Read);
            //Loads an existing Word document file stream
            using WordDocument wordDocument = new WordDocument(inputStream, Syncfusion.DocIO.FormatType.Docx);
            //Adds fallback font for "Arabic" custom unicode range
            wordDocument.FontSettings.FallbackFonts.Add(new FallbackFont(0x0600, 0x06ff, "Arial"));
            //Adds fallback font for "Hebrew" custom unicode range
            wordDocument.FontSettings.FallbackFonts.Add(new FallbackFont(0x0590, 0x05ff, "Times New Roman"));
            //Adds fallback font for "Hindi" custom unicode range
            wordDocument.FontSettings.FallbackFonts.Add(new FallbackFont(0x0900, 0x097F, "Nirmala UI"));
            //Adds fallback font for "Chinese" custom unicode range
            wordDocument.FontSettings.FallbackFonts.Add(new FallbackFont(0x4E00, 0x9FFF, "DengXian"));
            //Adds fallback font for "Japanese" custom unicode range
            wordDocument.FontSettings.FallbackFonts.Add(new FallbackFont(0x3040, 0x309F, "MS Gothic"));
            //Adds fallback font for "Thai" custom unicode range
            wordDocument.FontSettings.FallbackFonts.Add(new FallbackFont(0x0E00, 0x0E7F, "Tahoma"));
            //Adds fallback font for "Korean" custom unicode range
            wordDocument.FontSettings.FallbackFonts.Add(new FallbackFont(0xAC00, 0xD7A3, "Malgun Gothic"));
            //Instantiation of DocIORenderer for Word to Image conversion.
            using DocIORenderer render = new DocIORenderer();
            //Convert the entire Word document to images.
            Stream[] imageStreams = wordDocument.RenderAsImages();
            int i = 0;
            foreach (Stream stream in imageStreams)
            {
                //Reset the stream position.
                stream.Position = 0;
                //Save the stream as file.
                using FileStream fileStreamOutput = File.Create(@"../../../WordToImage_" + i + ".jpeg");
                stream.CopyTo(fileStreamOutput);
                i++;
            }
        }
    }
}
