using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Create_Docx_format_Word_document
{
    class Program
    {
        static void Main(string[] args)
        {
			//Creates a new instance of WordDocument (Empty Word Document).
            using (WordDocument document = new WordDocument())
            {
                //Adds a section and a paragraph to the document.
                document.EnsureMinimal();
                //Appends text to the last paragraph of the document.
                document.LastParagraph.AppendText("Hello World");
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
