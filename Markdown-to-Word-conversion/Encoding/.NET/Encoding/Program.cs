using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Convert_Word_to_Markdown
{
    class Program
    {
        static void Main(string[] args)
        {
            //Create a Word document instance.
            using (WordDocument document = new WordDocument())
            {
                //Set the encoding for the Markdown file.
                document.MdImportSettings.Encoding = System.Text.Encoding.UTF8;
                //Open the Markdown file.
                document.Open(Path.GetFullPath("Data/Input.md"));
                //Save as a Word document.
                document.Save(Path.GetFullPath(@"Output/Output.docx"), FormatType.Docx);
            }
        }
    }
}
