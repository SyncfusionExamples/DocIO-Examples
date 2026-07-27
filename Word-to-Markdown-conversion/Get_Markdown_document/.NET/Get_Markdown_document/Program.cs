using Syncfusion.DocIO.DLS;
using Syncfusion.Office.Markdown;
using System.IO;

namespace Convert_Word_to_Markdown
{
    class Program
    {
        static void Main(string[] args)
        {
            // Open an existing Word document.
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"Data/Input.docx")))
            {
                // Convert the Word document to Markdown.
                MarkdownDocument markdownDocument = document.GetMarkdownDocument();
                // Save or process the Markdown document as needed.
                markdownDocument.Save(Path.GetFullPath(@"Output/Output.md"));
                // Dispose the Markdown document.
                markdownDocument.Dispose();
            }
        }
    }
}
