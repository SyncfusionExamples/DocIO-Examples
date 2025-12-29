using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Extract_bookmark_Content
{
    class Program
    {
        static void Main(string[] args)
        {          
            // Create an input file stream to open the document
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open, FileAccess.Read))
            {
                //Creates a new Word document.
                using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
                {
                    // Create the bookmark navigator instance
                    BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(document);
                    // Move to the bookmark
                    bookmarkNavigator.MoveToBookmark("Adventure_Bkmk");
                    // Get the bookmark content as a new Word document part.
                    WordDocumentPart bookmarkPart = bookmarkNavigator.GetContent();
                    // Load the extracted content into a temporary Word document for modification or export.
                    using (WordDocument tempDoc = bookmarkPart.GetAsWordDocument())
                    {
                        //Creates file stream.
                        using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.html"), FileMode.Create, FileAccess.ReadWrite))
                        {
                            //Saves the Word document to file stream.
                            tempDoc.Save(outputFileStream, FormatType.Html);
                        }
                    }             
                }
            }
        }           
    }
}

