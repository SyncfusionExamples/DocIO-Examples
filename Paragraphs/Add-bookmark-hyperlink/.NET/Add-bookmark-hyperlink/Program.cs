using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Add_bookmark_hyperlink
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
                //Creates new Bookmark.
                paragraph.AppendBookmarkStart("Introduction");
                paragraph.AppendText("Hyperlink");
                paragraph.AppendBookmarkEnd("Introduction");
                paragraph.AppendText("\nA hyperlink is a reference or navigation element in a document to another section of the same document or to another document that may be on or part of a (different) domain.");
                paragraph = section.AddParagraph();
                paragraph.AppendText("Bookmark Hyperlink: ");
                paragraph = section.AddParagraph();
                //Appends Bookmark hyperlink to the paragraph.
                paragraph.AppendHyperlink("Introduction", "Bookmark", HyperlinkType.Bookmark);
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
