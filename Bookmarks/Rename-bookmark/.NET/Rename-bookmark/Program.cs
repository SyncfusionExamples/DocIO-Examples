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
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite))
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
        /// <param name="existingbookmarkName">The name of the bookmark to replace.</param>
        /// <param name="replaceBookmarkName">The new name for the bookmark.</param>

        private static void ReplaceBookmarkName(WordDocument document, string existingbookmarkName, string replaceBookmarkName)
        {
            //Gets the bookmark instance by using FindByName method of BookmarkCollection with bookmark name
            Bookmark bookmark = document.Bookmarks.FindByName(existingbookmarkName);
            // No bookmark found, return immediately
            if (bookmark == null)
                return;
            //Gets owner paragraph of the bookmark
            WParagraph bookmarkStartParagraph = bookmark.BookmarkStart.Owner as WParagraph;
            //Gets index of the bookmark start and end
            int boomarkStartIndex = bookmarkStartParagraph.ChildEntities.IndexOf(bookmark.BookmarkStart);
            int boomarkEndIndex = 0;
            //Gets bookmark end paragraph
            WParagraph bookmarkEndParagraph = bookmark.BookmarkEnd.Owner as WParagraph;
            //Checks whether the bookmark start and end is in same paragraph
            if (bookmarkEndParagraph == bookmarkStartParagraph)
                boomarkEndIndex = bookmarkStartParagraph.ChildEntities.IndexOf(bookmark.BookmarkEnd);
            else
                boomarkEndIndex = bookmarkEndParagraph.ChildEntities.IndexOf(bookmark.BookmarkEnd);
            //Removes the bookmark from Word document.
            document.Bookmarks.Remove(bookmark);
            //Inserts new bookmark in place of deleted bookmark
            bookmarkStartParagraph.ChildEntities.Insert(boomarkStartIndex, bookmarkStartParagraph.AppendBookmarkStart(replaceBookmarkName));
            //Inserts bookmark end in corresponding paragraph.
            if (bookmarkEndParagraph == bookmarkStartParagraph)
                bookmarkStartParagraph.ChildEntities.Insert(boomarkEndIndex, bookmarkStartParagraph.AppendBookmarkEnd(replaceBookmarkName));
            else
                bookmarkEndParagraph.ChildEntities.Insert(boomarkEndIndex, bookmarkEndParagraph.AppendBookmarkEnd(replaceBookmarkName));
        }
        #endregion       
    }
}