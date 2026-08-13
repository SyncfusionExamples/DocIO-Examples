using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office.Markdown;
using System.IO;

namespace Open_Word_Document_From_Markdown
{
    class Program
    {
        static void Main(string[] args)
        {
            // Opens an existing Markdown document.
            using (MarkdownDocument markdownDocument = new MarkdownDocument(Path.GetFullPath("Input.md")))
            {
                // Creates a new WordDocument instance.
                using (WordDocument wordDocument = new WordDocument())
                {
                    // Loads the Markdown document content into the Word document.
                    wordDocument.Open(markdownDocument);
                    // Saves the Word document as a DOCX file.
                    wordDocument.Save(Path.GetFullPath("Output.docx"), FormatType.Docx);
                }
            }
        }
    }
}
