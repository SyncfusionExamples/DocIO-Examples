using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;

namespace Insert_paragraph_item_into_bookmark
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Creates a new Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    //Creates the bookmark navigator instance to access the bookmark.
                    BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(document);
                    //Moves the virtual cursor to the location before the end of the bookmark "Northwind".
                    bookmarkNavigator.MoveToBookmark("Northwind", false, true);
                    //Inserts a new picture after the bookmark end.
                    WPicture picture = bookmarkNavigator.InsertParagraphItem(ParagraphItemType.Picture) as WPicture;
                    FileStream imageStream = new FileStream(Path.GetFullPath(@"../../../Data/Northwind.png"), FileMode.Open, FileAccess.Read);
                    picture.LoadImage(imageStream);
                    picture.WidthScale = 50;
                    picture.HeightScale = 50;
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
