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
                    ReplaceBookmarkText(document, "Description");
                    ReplaceBookmarkText(document, "Adventure");
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
            foreach (TextBodyItem item in textBodyPart.BodyItems)
            {
                IterateTextBody(item);
            }
            //Replace the bookmark content with text body part.
            bookmarksNavigator.ReplaceBookmarkContent(textBodyPart);
        }
        /// <summary>
        /// Iterate text body items.
        /// </summary>
        public static void IterateTextBody(TextBodyItem item)
        {
            switch (item.EntityType)
            {
                case EntityType.Paragraph:
                    WParagraph paragraph = (WParagraph)item;
                    //Replace a text in the bookmark content.
                    paragraph.Replace(new System.Text.RegularExpressions.Regex("Price"), "Amount");
                    paragraph.Replace(new System.Text.RegularExpressions.Regex("2000"), "two thousand");
                    break;
                case EntityType.Table:
                    WTable table = (WTable)item;
                    foreach (WTableRow row in table.Rows)
                    {
                        foreach (WTableCell cell in row.Cells)
                        {
                            foreach (TextBodyItem bodyitem in cell.ChildEntities)
                            {
                                IterateTextBody(bodyitem);
                            }
                        }
                    }
                    break;

            }
        }
    }
}
