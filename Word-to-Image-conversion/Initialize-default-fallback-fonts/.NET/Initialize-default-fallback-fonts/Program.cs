using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;

namespace Initialize_default_fallback_fonts
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Opens the file as stream.
            using FileStream inputStream = new FileStream(@"Data/Template.docx", FileMode.Open, FileAccess.Read);
            //Loads an existing Word document file stream.
            using WordDocument wordDocument = new WordDocument(inputStream, Syncfusion.DocIO.FormatType.Docx);
            //Initialize the default fallback fonts collection.
            wordDocument.FontSettings.FallbackFonts.InitializeDefault();
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
                using FileStream fileStreamOutput = File.Create(@"Output/Output_" + i + ".jpeg");
                stream.CopyTo(fileStreamOutput);
                i++;
            }
        }
    }
}
