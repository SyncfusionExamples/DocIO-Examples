using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Convert_Markdown_to_Word
{
    class Program
    {
        static void Main(string[] args)
        {
            // Open an existing Markdown file.
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"Data/Input.md")))
            {
                // Save as a Word document.
                document.Save(Path.GetFullPath(@"Output/MarkdownToWord.docx"), FormatType.Docx);
            }
        }
    }
}
