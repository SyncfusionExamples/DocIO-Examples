using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Preserve_bookmarks_from_cloned_content
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Open an existing Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    //Create the bookmark navigator instance to access the bookmark.
                    BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(document);
                    //Moves the virtual cursor to the location before the end of the bookmark "MainContent".
                    bookmarkNavigator.MoveToBookmark("MainContent");
                    //Get the content of the "MainContent" bookmark as WordDocumentPart.
                    WordDocumentPart wordDocumentPart = bookmarkNavigator.GetContent();

                    //Find the "MainContent" bookmark in the document by its name.
                    Bookmark bookmark = document.Bookmarks.FindByName("MainContent");
                    //Identify the parent text body of the bookmark.
                    WTextBody textbody = bookmark.BookmarkStart.OwnerParagraph.Owner as WTextBody;
                    //Determine the index of the paragraph containing the bookmark.
                    int index = textbody.ChildEntities.IndexOf(bookmark.BookmarkStart.OwnerParagraph);

                    //Remove the "MainContent" bookmark from the document.
                    document.Bookmarks.Remove(bookmark);
                    //Remove inner bookmarks (SubContent1, SubContent2, SubContent3) from the document.
                    bookmark = document.Bookmarks.FindByName("SubContent1");
                    document.Bookmarks.Remove(bookmark);
                    bookmark = document.Bookmarks.FindByName("SubContent2");
                    document.Bookmarks.Remove(bookmark);
                    bookmark = document.Bookmarks.FindByName("SubContent3");
                    document.Bookmarks.Remove(bookmark);

                    //Insert the cloned content of the "MainContent" bookmark after the original bookmark paragraph.
                    if (wordDocumentPart.Sections[0].ChildEntities[0] is WTextBody)
                    {
                        WTextBody clonedTextBody = wordDocumentPart.Sections[0].ChildEntities[0] as WTextBody;
                        for (int i = 0, j = index + 1; i < clonedTextBody.ChildEntities.Count; i++, j++)
                        {
                            textbody.ChildEntities.Insert(j, clonedTextBody.ChildEntities[i]);
                        }
                    }
                    //Close the WordDocumentPart instance.
                    wordDocumentPart.Close();
                    //Create file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
