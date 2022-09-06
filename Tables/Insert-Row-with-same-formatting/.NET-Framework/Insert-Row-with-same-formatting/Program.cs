using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;
using System.Text.RegularExpressions;

namespace Insert_Row_with_same_formatting
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open an existing Word document.
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"../../Data/Input.docx"), FormatType.Automatic))
            {
                //Get the table from the Word document.
                WTable table = document.Sections[0].Tables[0] as WTable;
                //Clone the row.
                WTableRow row = table.Rows[0].Clone();

                //Iterate all cells in row and clear the contents.
                for (int i = 0; i < row.Cells.Count; i++)
                {
                    WTableCell tableCell = row.Cells[i];
                    tableCell.ChildEntities.Clear();
                }
                //Insert new paragraph to the first cell.
                WParagraph cellParagraph = row.Cells[0].AddParagraph() as WParagraph;
                //Set text to the paragraph.
                IWTextRange textRange = cellParagraph.AppendText("New row's first cell");
                //Insert a row into the table in specific index.
                table.Rows.Insert(2, row);
                //Save the Word document.
                document.Save(Path.GetFullPath(@"../../Result.docx"), FormatType.Docx);
            }
        }
    }
}
