using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Replace_content_with_document_part
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an existing Word document.
                using (WordDocument templateDocument = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    //Creates the bookmark navigator instance to access the bookmark.
                    BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(templateDocument);
                    //Moves the virtual cursor to the location before the end of the bookmark "Northwind".
                    bookmarkNavigator.MoveToBookmark("Northwind");
                    //Gets the bookmark content as WordDocumentPart.
                    WordDocumentPart wordDocumentPart = bookmarkNavigator.GetContent();
                    //Loads the Word document with bookmark NorthwindDB.
                    using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Bookmarks.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
                        {
                            //Creates the bookmark navigator instance to access the bookmark.
                            bookmarkNavigator = new BookmarksNavigator(document);
                            //Moves the virtual cursor to the location before the end of the bookmark "NorthwindDB".
                            bookmarkNavigator.MoveToBookmark("NorthwindDB");
                            //Replaces the bookmark content with word body part.
                            bookmarkNavigator.ReplaceContent(wordDocumentPart);
                            //Close the WordDocumentPart instance.
                            wordDocumentPart.Close();
                            //Creates file stream.
                            using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite))
                            {
                                //Saves the Word document to file stream.
                                document.Save(outputFileStream, FormatType.Docx);
                            }
                        }
                    }
                }
            }
        }
    }
}
