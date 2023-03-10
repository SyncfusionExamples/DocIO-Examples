using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Replace_bookmark_content_with_another_document
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Data/DestinationWordDocument.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Open destination Word document.
                using (WordDocument destinationWordDocument = new WordDocument(fileStreamPath, FormatType.Automatic)) 
                {
                    using (FileStream sourceFileStream = new FileStream(Path.GetFullPath(@"../../../Data/SourceWordDocument.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        //Open the source Word document to copy all the content.
                        using (WordDocument sourceWordDocument = new WordDocument(sourceFileStream, FormatType.Automatic))
                        {
                            //Get all the content as Word document part.
                            WordDocumentPart wordDocumentPart = new WordDocumentPart(sourceWordDocument);
                            //Create the bookmark navigator instance to access the bookmark.
                            BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(destinationWordDocument);
                            //Move the virtual cursor to the location before the end of the bookmark "Adventure_Bkmk".
                            bookmarkNavigator.MoveToBookmark("Adventure_Bkmk");
                            //Replace the bookmark content with Word document part.
                            bookmarkNavigator.ReplaceContent(wordDocumentPart);
                            //Create file stream.
                            using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                            {
                                //Save the Word document to file stream.
                                destinationWordDocument.Save(outputFileStream, FormatType.Docx);
                            }
                        }
                    }
                }
            }
        }
    }
}
