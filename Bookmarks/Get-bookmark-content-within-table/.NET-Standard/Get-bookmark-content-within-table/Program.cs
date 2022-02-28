using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;

namespace Get_bookmark_content_within_table
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Adds a section and a paragraph in the document.
                document.EnsureMinimal();
                //Inserts a new table with bookmark.
                IWTable table = CreateTable(document);
                //Creates the bookmark navigator instance to access the bookmark.
                BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(document);
                //Moves the virtual cursor to the location before the end of the bookmark "BkmkInTable".
                bookmarkNavigator.MoveToBookmark("BkmkInTable");
                //Sets the column index where the bookmark starts within the table.
                bookmarkNavigator.CurrentBookmark.FirstColumn = 1;
                //Sets the column index where the bookmark ends within the table.
                bookmarkNavigator.CurrentBookmark.LastColumn = 4;
                //Gets the bookmark content.
                TextBodyPart part = bookmarkNavigator.GetBookmarkContent();
                //Adds new section.
                document.AddSection();
                for (int i = 0; i < part.BodyItems.Count; i++)
                    //Adds the retrieved content into another new section.
                    document.LastSection.Body.ChildEntities.Add(part.BodyItems[i]);
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }

        /// <summary>
        /// Creates the table.
        /// </summary>
        /// <returns></returns>
        public static IWTable CreateTable(WordDocument document)
        {
            //Adds a new table into Word document.
            IWTable table = document.LastSection.AddTable();
            //Specifies the total number of rows & columns.
            table.ResetCells(5, 5);
            //Accesses the instance of the cells and adds the content into cells.
            table[0, 0].AddParagraph().AppendText("Supplier ID");
            table[0, 1].AddParagraph().AppendText("Company Name");
            IWParagraph paragraph = table.Rows[0].Cells[2].AddParagraph();
            //Appends a bookmark start in third cell of first row.
            paragraph.AppendBookmarkStart("BkmkInTable");
            paragraph.AppendText("Contact Name");
            table[0, 3].AddParagraph().AppendText("Address");
            table[0, 4].AddParagraph().AppendText("City");
            table[1, 0].AddParagraph().AppendText("1");
            table[1, 1].AddParagraph().AppendText("Exotic Liquids");
            table[1, 2].AddParagraph().AppendText("Charlotte Cooper");
            table[1, 3].AddParagraph().AppendText("49 Gilbert St.");
            table[1, 4].AddParagraph().AppendText("London");
            table[2, 0].AddParagraph().AppendText("2");
            table[2, 1].AddParagraph().AppendText("New Orleans Cajun Delights");
            table[2, 2].AddParagraph().AppendText("Shelley Burke");
            table[2, 3].AddParagraph().AppendText("P.O. Box 78934");
            table[2, 4].AddParagraph().AppendText("New Orleans");
            table[3, 0].AddParagraph().AppendText("3");
            table[3, 1].AddParagraph().AppendText("Grandma Kelly's Homestead");
            table[3, 2].AddParagraph().AppendText("Regina Murphy");
            table[3, 3].AddParagraph().AppendText("707 Oxford Rd.");
            table[3, 4].AddParagraph().AppendText("Ann Arbor");
            table[4, 0].AddParagraph().AppendText("4");
            table[4, 1].AddParagraph().AppendText("Tokyo Traders");
            paragraph = table.Rows[4].Cells[2].AddParagraph();
            //Appends a bookmark end in third cell of last row.
            paragraph.AppendBookmarkEnd("BkmkInTable");
            paragraph.AppendText("Yoshi Nagase");
            table[4, 3].AddParagraph().AppendText("9-8 Sekimai Musashino - shi");
            table[4, 4].AddParagraph().AppendText("Tokyo");
            return table;
        }
    }
}
