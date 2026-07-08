using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Convert_Word_to_Markdown
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open an existing Word document.
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"Data/Input.docx")))
            {
                //Save the document as a Markdown file.
                document.Save(Path.GetFullPath(@"Output/Output.md"));
            }      
        }
    }
}
