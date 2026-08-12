using Syncfusion.DocIO.DLS;
using System.IO;
using System.Text;

namespace Convert_Word_To_Markdown_with_Encoding
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open an existing Word document.
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"Data/Input.docx")))
            {
                //Set the encoding values.
                document.SaveOptions.MarkdownSaveOptions.Encoding = Encoding.ASCII;
                //Save the document as a Markdown file.
                document.Save(Path.GetFullPath(@"Output/Output.md"));
            }
        }
    }
}
