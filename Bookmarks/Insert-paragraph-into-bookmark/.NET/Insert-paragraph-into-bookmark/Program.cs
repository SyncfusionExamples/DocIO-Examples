using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;

namespace Insert_paragraph_into_bookmark
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
                    //Creates the bookmark navigator instance to access the bookmark.
                    BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(document);
                    //Moves the virtual cursor to the location before the end of the bookmark "Northwind".
                    bookmarkNavigator.MoveToBookmark("Northwind", false, true);
                    //Inserts a new paragraph before the bookmark start.
                    IWParagraph paragraph = new WParagraph(document);
                    paragraph.AppendText("Northwind Database is a set of tables containing data fitted into predefined categories.");
                    bookmarkNavigator.InsertParagraph(paragraph);
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
