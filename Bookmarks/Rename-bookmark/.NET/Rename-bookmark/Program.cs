using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Rename_bookmark
{
    class Program
    {
        static void Main(string[] args)
        {       
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an existing Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    //Replace Bookmark name
                    ReplaceBookmarkName(document, "Northwind", "New_Bookmark");
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
        #region Replace Bookmark name
        /// <summary>
        /// Replace bookmark name
        /// </summary>
        /// <param name="document">Input Word document.</param>
        /// <param name="existingBookmarkName">The name of the bookmark to replace.</param>
        /// <param name="replaceBookmarkName">The new name for the bookmark.</param>
        private static void ReplaceBookmarkName(WordDocument document, string existingBookmarkName, string replaceBookmarkName)
        {
            //Gets the bookmark instance by using FindByName method of BookmarkCollection with bookmark name
            Bookmark bookmark = document.Bookmarks.FindByName(existingBookmarkName);
            // No bookmark found, return immediately
            if (bookmark == null)
                return;
            // Variables for owner entity start and end positions
            WParagraph startParagraph = null;
            InlineContentControl startICC = null;
            int startIndex = -1;

            WParagraph endParagraph = null;
            InlineContentControl endICC = null;
            int endIndex = -1;
            // Determine the owner and index for the bookmark start.
            // The bookmark start may be inside a WParagraph (as a child entity)
            // or inside an InlineContentControl (as a paragraph item).
            if (bookmark.BookmarkStart != null && bookmark.BookmarkStart.Owner is WParagraph)
            {
                startParagraph = bookmark.BookmarkStart.Owner as WParagraph;
                startIndex = startParagraph.ChildEntities.IndexOf(bookmark.BookmarkStart);
            }
            else if (bookmark.BookmarkStart != null && bookmark.BookmarkStart.Owner is InlineContentControl)
            {
                startICC = bookmark.BookmarkStart.Owner as InlineContentControl;
                startIndex = startICC.ParagraphItems.IndexOf(bookmark.BookmarkStart);
            }
            // Determine the owner and index for the bookmark end.
            // Similar to start, the end could be in a paragraph or inline content contro
            if (bookmark.BookmarkEnd != null && bookmark.BookmarkEnd.Owner is WParagraph)
            {
                endParagraph = bookmark.BookmarkEnd.Owner as WParagraph;
                endIndex = endParagraph.ChildEntities.IndexOf(bookmark.BookmarkEnd);
            }
            else if (bookmark.BookmarkEnd != null && bookmark.BookmarkEnd.Owner is InlineContentControl)
            {
                endICC = bookmark.BookmarkEnd.Owner as InlineContentControl;
                endIndex = endICC.ParagraphItems.IndexOf(bookmark.BookmarkEnd);
            }
            //Removes the bookmark from Word document.
            document.Bookmarks.Remove(bookmark);
            // Create a new BookmarkStart and insert at the recorded index.
            BookmarkStart newBookmarkStart = new BookmarkStart(document, replaceBookmarkName);
            // Insert new bookmark start at the original position with the new name.
            if (startParagraph != null)
                startParagraph.ChildEntities.Insert(startIndex, newBookmarkStart);
            else if (startICC != null)
                startICC.ParagraphItems.Insert(startIndex, newBookmarkStart);
            // Create a new BookmarkEnd and insert at the recorded index.
            BookmarkEnd newBookmarkEnd = new BookmarkEnd(document, replaceBookmarkName);
            // Insert new bookmark end at the original position with the new name.
            if (endParagraph != null)
                endParagraph.ChildEntities.Insert(endIndex, newBookmarkEnd);
            else if (endICC != null)           
                endICC.ParagraphItems.Insert(endIndex, newBookmarkEnd);
        }
        #endregion       
    }
}