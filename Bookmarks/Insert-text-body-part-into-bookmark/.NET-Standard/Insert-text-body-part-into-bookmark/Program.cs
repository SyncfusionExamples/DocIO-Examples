using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;

namespace Insert_text_body_part_into_bookmark
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an existing Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    //Creates the bookmark navigator instance to access the bookmark.
                    BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(document);
                    //Moves the virtual cursor to the location before the end of the bookmark "Northwind".
                    bookmarkNavigator.MoveToBookmark("Northwind");
                    //Gets the bookmark content.
                    TextBodyPart textBodyPart = bookmarkNavigator.GetBookmarkContent();
                    document.AddSection();
                    IWParagraph paragraph = document.LastSection.AddParagraph();
                    paragraph.AppendText("Northwind Database is a set of tables containing data fitted into predefined categories.");
                    //Adds the new bookmark into Word document.
                    paragraph.AppendBookmarkStart("bookmark_empty");
                    paragraph.AppendBookmarkEnd("bookmark_empty");
                    //Moves the virtual cursor to the location after the start of the bookmark "bookmark_empty".
                    bookmarkNavigator.MoveToBookmark("bookmark_empty", true, true);
                    //Inserts the text body part after the bookmark start.
                    bookmarkNavigator.InsertTextBodyPart(textBodyPart);
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
}
