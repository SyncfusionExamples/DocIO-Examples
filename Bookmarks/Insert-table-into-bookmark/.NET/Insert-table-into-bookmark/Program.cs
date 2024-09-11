using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;

namespace Insert_table_into_bookmark
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
                    bookmarkNavigator.MoveToBookmark("Northwind", false, false);
                    //Inserts a new paragraph before the bookmark end.
                    IWParagraph paragraph = new WParagraph(document);
                    paragraph.AppendText("Northwind Database Contains the following tables:");
                    bookmarkNavigator.InsertParagraph(paragraph);
                    //Inserts a new table before the bookmark end.
                    WTable table = new WTable(document);
                    table.ResetCells(3, 2);
                    table[0, 0].AddParagraph().AppendText("Suppliers");
                    table[0, 1].AddParagraph().AppendText("2");
                    table[1, 0].AddParagraph().AppendText("Customers");
                    table[1, 1].AddParagraph().AppendText("1");
                    table[2, 0].AddParagraph().AppendText("Employees");
                    table[2, 1].AddParagraph().AppendText("3");
                    bookmarkNavigator.InsertTable(table);
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
