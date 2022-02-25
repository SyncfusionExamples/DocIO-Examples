using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;

namespace Split_a_document_by_bookmark
{
    class Program
    {
        static void Main(string[] args)
        {

            //Load an existing Word document.
            FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
            {
                //Create the bookmark navigator instance to access the bookmark.
                BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(document);
                //Get the bookmark collections in the document.
                BookmarkCollection bookmarkCollection =  document.Bookmarks;
                foreach (Bookmark bookmark in bookmarkCollection)
                {
                    //Move the virtual cursor to the location before the end of the bookmark.
                    bookmarkNavigator.MoveToBookmark(bookmark.Name);
                    //Get the bookmark content.
                    TextBodyPart part = bookmarkNavigator.GetBookmarkContent();
                    //Create a new Word document.
                    WordDocument newDocument = new WordDocument();
                    newDocument.AddSection();
                    //Add the retrieved content into another new document.
                    for (int i = 0; i < part.BodyItems.Count; i++)
                        newDocument.LastSection.Body.ChildEntities.Add(part.BodyItems[i].Clone());
                    //Save the Word document to file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result"+ bookmark.Name + ".docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        newDocument.Save(outputFileStream, FormatType.Docx);
                    }
                }
                

            }
            
        }
    }
}
