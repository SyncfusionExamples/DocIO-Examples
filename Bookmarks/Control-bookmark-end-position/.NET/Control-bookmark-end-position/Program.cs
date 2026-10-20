using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Control_bookmark_end_position
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates an instance of WordDocument class (Empty Word Document).
            using (WordDocument document = new WordDocument())
            {
                //Adds a section to the document.
                IWSection section = document.AddSection();
                //Adds a paragraph and preserves the bookmark end after the paragraph mark.
                IWParagraph paragraph = section.AddParagraph();
                paragraph.AppendBookmarkStart("ParagraphBookmark");
                paragraph.AppendText("Paragraph bookmark");
                BookmarkEnd bookmarkEnd = paragraph.AppendBookmarkEnd("ParagraphBookmark");
                bookmarkEnd.IsAfterParagraphMark = true;
                //Adds a 2 x 2 table.
                IWTable table = section.AddTable();
                table.ResetCells(2, 2);
                //Adds a bookmark end after the table-cell mark of the first cell.
                IWParagraph cellParagraph = table[0, 0].AddParagraph();
                cellParagraph.AppendBookmarkStart("CellBookmark");
                cellParagraph.AppendText("Cell bookmark");
                bookmarkEnd = cellParagraph.AppendBookmarkEnd("CellBookmark");
                bookmarkEnd.IsAfterCellMark = true;
                //Adds a bookmark end after the first table-row mark, in the final cell of that row.
                IWParagraph rowParagraph = table[0, 1].AddParagraph();
                rowParagraph.AppendBookmarkStart("RowBookmark");
                rowParagraph.AppendText("Row bookmark");
                bookmarkEnd = rowParagraph.AppendBookmarkEnd("RowBookmark");
                bookmarkEnd.IsAfterRowMark = true;
                //Adds content to the remaining cell without a bookmark end.
                table[1, 0].AddParagraph().AppendText("Content");
                //Adds a bookmark end after the table mark, in the final cell of the final row.
                IWParagraph tableParagraph = table[1, 1].AddParagraph();
                tableParagraph.AppendBookmarkStart("TableBookmark");
                tableParagraph.AppendText("Table bookmark");
                bookmarkEnd = tableParagraph.AppendBookmarkEnd("TableBookmark");
                bookmarkEnd.IsAfterTableMark = true;
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/BookmarkEndPositions.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
