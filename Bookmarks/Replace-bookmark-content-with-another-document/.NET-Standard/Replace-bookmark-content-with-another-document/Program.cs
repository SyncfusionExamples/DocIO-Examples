using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Replace_bookmark_content_with_another_document
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Open destination Word document.
                using (WordDocument templateDocument = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    using (FileStream sourceFileStream = new FileStream(Path.GetFullPath(@"../../../Data/variation.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        //Open an source Word document for copying all the content.
                        using (WordDocument sourceDocument = new WordDocument(sourceFileStream, FormatType.Automatic))
                        {
                            //Get all the content as Word document part.
                            WordDocumentPart wordDocumentPart = new WordDocumentPart(sourceDocument);
                            //Create the bookmark navigator instance to access the bookmark.
                            BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(templateDocument);
                            //Move the virtual cursor to the location before the end of the bookmark "Adventure_Bkmk".
                            bookmarkNavigator.MoveToBookmark("Adventure_Bkmk");
                            //Replace the bookmark content with Word document part.
                            bookmarkNavigator.ReplaceContent(wordDocumentPart);
                            //Create file stream.
                            using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                            {
                                //Save the Word document to file stream.
                                templateDocument.Save(outputFileStream, FormatType.Docx);
                            }
                        }
                    }
                }
            }
        }
    }
}
