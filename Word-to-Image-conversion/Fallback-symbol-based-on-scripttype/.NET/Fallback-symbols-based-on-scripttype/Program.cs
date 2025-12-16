using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Office;

namespace Fallback_symbols_based_on_scripttype
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Opens the file as stream.
            using FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read);
            //Loads an existing Word document file stream.
            using WordDocument wordDocument = new WordDocument(inputStream, Syncfusion.DocIO.FormatType.Docx);
            //Adds fallback font for basic symbols like bullet characters.
            wordDocument.FontSettings.FallbackFonts.Add(ScriptType.Symbols, "Segoe UI Symbol, Arial Unicode MS, Wingdings");
            //Adds fallback font for mathematics symbols.
            wordDocument.FontSettings.FallbackFonts.Add(ScriptType.Mathematics, "Cambria Math, Noto Sans Math, Segoe UI Symbol, Arial Unicode MS");
            //Adds fallback font for emojis.
            wordDocument.FontSettings.FallbackFonts.Add(ScriptType.Emoji, "Segoe UI Emoji, Noto Color Emoji, Arial Unicode MS");
            //Instantiation of DocIORenderer for Word to image conversion.
            using DocIORenderer render = new DocIORenderer();
            //Convert the entire Word document to images.
            Stream[] imageStreams = wordDocument.RenderAsImages();
            int i = 0;
            foreach (Stream stream in imageStreams)
            {
                //Reset the stream position.
                stream.Position = 0;
                //Save the stream as file.
                using FileStream fileStreamOutput = File.Create(Path.GetFullPath(@"../../../Output/Output_" + i + ".jpeg"));
                stream.CopyTo(fileStreamOutput);
                i++;
            }
        }
    }
}
