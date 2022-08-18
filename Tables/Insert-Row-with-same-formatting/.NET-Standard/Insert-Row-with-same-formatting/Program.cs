using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Insert_Row_with_same_formatting
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Data/Input.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Opens an existing Word document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Automatic))
                {
                    //Gets the table from the Word document
                    WTable table = document.Sections[0].Tables[0] as WTable;
                    //Clones the row, from the similar type of row
                    WTableRow row = table.Rows[1].Clone();

                    //Clears existing content from cloned row.
                    for (int i = 0; i < row.Cells.Count; i++)
                    {
                        //Gets the cells
                        WTableCell tableCell = row.Cells[i];
                        //Clears cell child entites.
                        tableCell.ChildEntities.Clear();
                    }
                    //Insert new paragraph to the first cell
                    WParagraph cellParagraph = row.Cells[0].AddParagraph() as WParagraph;
                    //Set text and character format.
                    IWTextRange textRange = cellParagraph.AppendText("Hello World");
                    textRange.CharacterFormat.Bold = true;
                    //Insert new row into the table in specific index.
                    table.Rows.Insert(2, row);
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
