using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Create_nested_table
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                IWSection section = document.AddSection();
                section.AddParagraph().AppendText("Price Details");
                IWTable table = section.AddTable();
                table.ResetCells(3, 2);
                table[0, 0].AddParagraph().AppendText("Item");
                table[0, 1].AddParagraph().AppendText("Price($)");
                table[1, 0].AddParagraph().AppendText("Items with same price");
                //Adds a nested table into the cell (second row, first cell).
                IWTable nestTable = table[1, 0].AddTable();
                //Creates the specified number of rows and columns to nested table.
                nestTable.ResetCells(3, 1);
                //Accesses the instance of the nested table cell (first row, first cell).
                WTableCell nestedCell = nestTable.Rows[0].Cells[0];
                //Specifies the width of the nested cell.
                nestedCell.Width = 200;
                //Adds the content into nested cell.
                nestedCell.AddParagraph().AppendText("Apple");
                //Accesses the instance of the nested table cell (second row, first cell).
                nestedCell = nestTable.Rows[1].Cells[0];
                //Specifies the width of the nested cell.
                nestedCell.Width = 200;
                //Adds the content into nested cell.
                nestedCell.AddParagraph().AppendText("Orange");
                //Accesses the instance of the nested table cell (third row, first cell).
                nestedCell = nestTable.Rows[2].Cells[0];
                //Specifies the width of the nested cell.
                nestedCell.Width = 200;
                //Adds the content into nested cell.
                nestedCell.AddParagraph().AppendText("Mango");
                //Accesses the instance of the cell (second row, second cell).
                nestedCell = table.Rows[1].Cells[1];
                table[1, 1].AddParagraph().AppendText("85");
                table[2, 0].AddParagraph().AppendText("Pomegranate");
                table[2, 1].AddParagraph().AppendText("70");
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
