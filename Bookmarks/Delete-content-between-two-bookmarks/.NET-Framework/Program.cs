using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace Delete_content_between_two_bookmarks
{
    internal class Program
    {
        static void Main(string[] args)
        {
            using (WordDocument document = new WordDocument(@"../../Data/Template.docx", FormatType.Docx))
            {
                //Delete the text between two different existing bookmarks. 
                DeleteBookmarkContent(document, "BM1", "BM2");
                //Save and close the Word document.
                document.Save(@"../../Data/Result.docx", FormatType.Docx);
            }
        }
        /// <summary>
        /// Deletes the content of a bookmark located between two other bookmarks in the Word document.
        /// A temporary bookmark is created between the two specified bookmarks to identify the content to delete.
        /// </summary>
        /// <param name="document">The <see cref="WordDocument"/> from which the bookmark content will be deleted.</param>
        /// <param name="bookmark1">The name of the first bookmark used to define the starting point of the content to delete.</param>
        /// <param name="bookmark2">The name of the second bookmark used to define the ending point of the content to delete.</param>
        public static void DeleteBookmarkContent(WordDocument document, String bookmark1, String bookmark2)
        {
            //Temp Bookmark.
            String tempBookmarkName = "tempBookmark";

            #region Insert bookmark start after bookmark1.
            //Get the bookmark instance by using FindByName method of BookmarkCollection with bookmark name.
            Bookmark firstBookmark = document.Bookmarks.FindByName(bookmark1);
            //Access the bookmark end’s owner paragraph by using bookmark.
            WParagraph firstBookmarkOwnerPara = firstBookmark.BookmarkEnd.OwnerParagraph;
            //Get the index of bookmark end of bookmark1.
            int index = firstBookmarkOwnerPara.Items.IndexOf(firstBookmark.BookmarkEnd);
            //Create and add new bookmark start after bookmark1.
            BookmarkStart newBookmarkStart = new BookmarkStart(document, tempBookmarkName);
            firstBookmarkOwnerPara.ChildEntities.Insert(index + 1, newBookmarkStart);
            #endregion

            #region Insert bookmark end before bookmark2.
            //Get the bookmark instance by using FindByName method of BookmarkCollection with bookmark name.
            Bookmark secondBookmark = document.Bookmarks.FindByName(bookmark2);
            //Access the bookmark start’s owner paragraph by using bookmark.
            WParagraph secondBookmarkOwnerPara = secondBookmark.BookmarkStart.OwnerParagraph;
            //Get the index of bookmark start of bookmark2.
            index = secondBookmarkOwnerPara.Items.IndexOf(secondBookmark.BookmarkStart);
            //Create and add new bookmark end before bookmark2.
            BookmarkEnd newBookmarkEnd = new BookmarkEnd(document, tempBookmarkName);
            secondBookmarkOwnerPara.ChildEntities.Insert(index, newBookmarkEnd);
            #endregion

            #region Select bookmark content and delete.
            //Create the bookmark navigator instance to access the newly created bookmark.
            BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(document);
            //Move the virtual cursor to the location of the temp bookmark.
            bookmarkNavigator.MoveToBookmark(tempBookmarkName);
            //Save bookmark content as Word document.
            ExportBookmarkContentToDocument(bookmarkNavigator);
            //Replace the bookmark content.
            bookmarkNavigator.DeleteBookmarkContent(false);
            #endregion

            #region Remove that temporary bookmark.
            //Get the bookmark instance by using FindByName method of BookmarkCollection with bookmark name.
            Bookmark bookmark = document.Bookmarks.FindByName(tempBookmarkName);
            //Remove the temp bookmark named from Word document.
            document.Bookmarks.Remove(bookmark);
            #endregion
        }
        /// <summary>
        /// Extracts the content of the current bookmark and saves it as a separate Word document (.docx).
        /// </summary>
        /// <param name="bookmarkNavigator">The <see cref="BookmarksNavigator"/> instance used to navigate and access the bookmark content.</param>
        private static void ExportBookmarkContentToDocument(BookmarksNavigator bookmarkNavigator)
        {
            //Get the bookmark content as WordDocumentPart.
            WordDocumentPart documentPart = bookmarkNavigator.GetContent();
            //Save the WordDocumentPart as separate Word document
            using (WordDocument newDocument = documentPart.GetAsWordDocument())
            {
                string tempBookmarkName = bookmarkNavigator.CurrentBookmark.Name;
                //Save the Word document to file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath($@"../../Data/{tempBookmarkName}.docx"), FileMode.Create, FileAccess.Write))
                {
                    newDocument.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
