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
            // Variables to store the index positions of the bookmark start and end within their respective owners
            int startIndex = -1;
            int endIndex = -1;
            // Create new bookmark start and end markers with the replacement name
            BookmarkStart newBookmarkStart = new BookmarkStart(document, replaceBookmarkName);
            BookmarkEnd newBookmarkEnd = new BookmarkEnd(document, replaceBookmarkName);
         
            // Determine the owner and index for the bookmark start.
            // The bookmark start may be inside a WParagraph (as a child entity)
            // or inside an InlineContentControl (as a paragraph item).
            if (bookmark.BookmarkStart != null && bookmark.BookmarkStart.Owner is WParagraph)
            {
                WParagraph startParagraph = bookmark.BookmarkStart.Owner as WParagraph;
                // Find the index of the old bookmark start in the paragraph's child entities
                startIndex = startParagraph.ChildEntities.IndexOf(bookmark.BookmarkStart);
                // Insert the new bookmark end at the same index.
                startParagraph.ChildEntities.Insert(startIndex, newBookmarkStart);
            }
            else if (bookmark.BookmarkStart != null && bookmark.BookmarkStart.Owner is InlineContentControl)
            {
                InlineContentControl startICC = bookmark.BookmarkStart.Owner as InlineContentControl;
                // Find the index of the old bookmark end in the ICC's paragraph items
                startIndex = startICC.ParagraphItems.IndexOf(bookmark.BookmarkStart);
                // Insert the new bookmark end at the same index.
                startICC.ParagraphItems.Insert(startIndex, newBookmarkStart);
            }
            // Determine the owner and index for the bookmark end.
            // Similar to start, the end could be in a paragraph or inline content control.
            if (bookmark.BookmarkEnd != null && bookmark.BookmarkEnd.Owner is WParagraph)
            {
                WParagraph endParagraph = bookmark.BookmarkEnd.Owner as WParagraph;
                // Find the index of the old bookmark end in the paragraph's child entities
                endIndex = endParagraph.ChildEntities.IndexOf(bookmark.BookmarkEnd);
                // Insert the new bookmark end at the same index.
                endParagraph.ChildEntities.Insert(endIndex, newBookmarkEnd);
            }
            else if (bookmark.BookmarkEnd != null && bookmark.BookmarkEnd.Owner is InlineContentControl)
            {
                InlineContentControl endICC = bookmark.BookmarkEnd.Owner as InlineContentControl;
                // Find the index of the old bookmark end in the ICC's paragraph items
                endIndex = endICC.ParagraphItems.IndexOf(bookmark.BookmarkEnd);
                // Insert the new bookmark end at the same index.
                endICC.ParagraphItems.Insert(endIndex, newBookmarkEnd);
            }
            //Removes the bookmark from Word document.
            document.Bookmarks.Remove(bookmark);        
        }
        #endregion       
    }
}