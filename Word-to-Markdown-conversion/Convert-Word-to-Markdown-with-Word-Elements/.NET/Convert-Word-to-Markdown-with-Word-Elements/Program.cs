using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using System.IO;

namespace Convert_Word_to_Markdown_with_Word_Elements
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open an existing Word document.
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"Data/Input.docx")))
            {
                // Initialize the DocIORenderer to preserve Word elements as images.
                using DocIORenderer docIORenderer = new DocIORenderer();
                //Save the document as a Markdown file.
                document.Save(Path.GetFullPath(@"Output/Output.md"));
            }
        }
    }
}

