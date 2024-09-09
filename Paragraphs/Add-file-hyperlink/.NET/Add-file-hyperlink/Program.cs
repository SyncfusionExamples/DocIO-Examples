using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Add_file_hyperlink
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Adds new section to the document.
                IWSection section = document.AddSection();
                //Adds new paragraph to the section.
                IWParagraph paragraph = section.AddParagraph();
                paragraph.AppendText("File Hyperlinks: ");
                paragraph = section.AddParagraph();
                //Appends hyperlink field to the paragraph.
                paragraph.AppendHyperlink(@"Data/Template.docx", "File", HyperlinkType.FileLink);
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
