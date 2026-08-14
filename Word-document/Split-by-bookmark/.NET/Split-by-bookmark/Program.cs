using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Compression.Zip;
using System.IO;

namespace Split_a_document_by_bookmark
{
    class Program
    {
        static void Main(string[] args)
        {
            //Load an existing Word document.
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (ZipArchive zipArchive = new ZipArchive())
                {


                    using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                    {
                        //Create the bookmark navigator instance to access the bookmark.
                        BookmarksNavigator bookmarksNavigator = new BookmarksNavigator(document);
                        BookmarkCollection bookmarkCollection = document.Bookmarks;
                        //Iterate each bookmark in Word document.
                        foreach (Bookmark bookmark in bookmarkCollection)
                        {
                            //Move the virtual cursor to the bookmark.
                            bookmarksNavigator.MoveToBookmark(bookmark.Name);
                            //Get the bookmark content as WordDocumentPart.
                            WordDocumentPart documentPart = bookmarksNavigator.GetContent();
                            //Save the WordDocumentPart as separate Word document
                            using (WordDocument newDocument = documentPart.GetAsWordDocument())
                            {
                                //Save the Word document to MemoryStream.
                                MemoryStream memoryStream = new MemoryStream();
                                newDocument.Save(memoryStream, FormatType.Docx);
                                //Add the Word document to Zip archive.
                                zipArchive.AddItem(bookmark.Name + ".docx", memoryStream, true, Syncfusion.Compression.FileAttributes.Normal);
                            }
                        }

                        zipArchive.Save(Path.GetFullPath(@"Output/Split-by-Bookmark.zip"));
                    }
                }

            }
        }
    }
}
