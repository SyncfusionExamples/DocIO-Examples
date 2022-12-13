using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;

namespace Replace_text_within_bookmark_content
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the file as a stream.
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"../../../Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Load the file stream into a Word document.
                using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
                {
                    //Replace a text within the bookmark.
                    ReplaceBookmarkText(document, "Adventure");
                    ReplaceBookmarkText(document, "Test");
                    //Create a file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to the file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
        public static void ReplaceBookmarkText(WordDocument document, string bookmarkName)
        {
            //Check whether the bookmark name is valid.
            if (string.IsNullOrEmpty(bookmarkName) || document.Bookmarks.FindByName(bookmarkName) == null)
                return;
            //Move to the virtual cursor before the bookmark end location of the bookmark.
            BookmarksNavigator bookmarksNavigator = new BookmarksNavigator(document);
            bookmarksNavigator.MoveToBookmark(bookmarkName);
            //Replace the bookmark content with new text.
            TextBodyPart textBodyPart = bookmarksNavigator.GetBookmarkContent();
            //Get paragraph from the textBody part.
            foreach (WParagraph paragraph in textBodyPart.BodyItems)
            {
                //Replace a text in the bookmark content.
                paragraph.Replace(new System.Text.RegularExpressions.Regex("two thousand"), "2000");
                paragraph.Replace(new System.Text.RegularExpressions.Regex("Washington"), "USA");
            }
            //Replace the bookmark content with text body part.
            bookmarksNavigator.ReplaceBookmarkContent(textBodyPart);
        }
    }
}
