using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;

namespace Get_bookmark_content_as_document_part
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Creates a new Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    //Creates the bookmark navigator instance to access the bookmark.
                    BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(document);
                    //Moves the virtual cursor to the location before the end of the bookmark "Northwind".
                    bookmarkNavigator.MoveToBookmark("Northwind");
                    //Gets the bookmark content as WordDocumentPart.
                    WordDocumentPart wordDocumentPart = bookmarkNavigator.GetContent();
                    //Saves the WordDocumentPart as separate Word document.
                    using (WordDocument newDocument = wordDocumentPart.GetAsWordDocument())
                    {
                        //Close the WordDocumentPart instance.
                        wordDocumentPart.Close();
                        //Creates file stream.
                        using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                        {
                            //Saves the Word document to file stream.
                            newDocument.Save(outputFileStream, FormatType.Docx);
                        }
                    }
                }
            }
        }
    }
}
