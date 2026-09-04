using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using System.IO;

namespace Embedding_Fonts
{
    class Program
    {
        static void Main(string[] args)
        {
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"Data/Input.docx")))
            {
                //Initialize the DocIORenderer.
                using DocIORenderer renderer = new DocIORenderer();
                //Enable the flag to embed complete TrueType/OpenType fonts used in the document.
                document.SaveOptions.EmbedFonts = true;
                //Save the Word document
                document.Save(Path.GetFullPath(@"Output/Output.md"));
                document.Close();
            }
        }
    }
}