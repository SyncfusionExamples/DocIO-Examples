using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Office;
namespace Fallback_fonts_based_on_scripttype
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Opens the file as stream
            using FileStream inputStream = new FileStream(@"../../../Template.docx", FileMode.Open, FileAccess.Read);
            //Loads an existing Word document file stream
            using WordDocument wordDocument = new WordDocument(inputStream, Syncfusion.DocIO.FormatType.Docx);
            //Adds fallback font for "Arabic" script type
            wordDocument.FontSettings.FallbackFonts.Add(ScriptType.Arabic, "Arial, Times New Roman");
            //Adds fallback font for "Hebrew" script type
            wordDocument.FontSettings.FallbackFonts.Add(ScriptType.Hebrew, "Arial, Courier New");
            //Adds fallback font for "Hindi" script type
            wordDocument.FontSettings.FallbackFonts.Add(ScriptType.Hindi, "Mangal, Nirmala UI");
            //Adds fallback font for "Chinese" script type
            wordDocument.FontSettings.FallbackFonts.Add(ScriptType.Chinese, "DengXian, MingLiU");
            //Adds fallback font for "Japanese" script type
            wordDocument.FontSettings.FallbackFonts.Add(ScriptType.Japanese, "Yu Mincho, MS Mincho");
            //Adds fallback font for "Thai" script type
            wordDocument.FontSettings.FallbackFonts.Add(ScriptType.Thai, "Tahoma, Microsoft Sans Serif");
            //Adds fallback font for "Korean" script type
            wordDocument.FontSettings.FallbackFonts.Add(ScriptType.Korean, "Malgun Gothic, Batang");
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
