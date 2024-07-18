using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Convert_Word_to_Markdown
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open a file as a stream.
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Load the file stream into a Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    //Create a file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../WordToMd.md"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save a Markdown file to the file stream.
                        document.Save(outputFileStream, FormatType.Markdown);
                    }
                }
            }
        }
    }
}
