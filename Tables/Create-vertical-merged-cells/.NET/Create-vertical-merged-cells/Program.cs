using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Create_vertical_merged_cells
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                IWSection section = document.AddSection();
                section.AddParagraph().AppendText("Vertical merging of Table cells");
                IWTable table = section.AddTable();
                table.ResetCells(2, 2);
                //Adds content to table cells.
                table[0, 0].AddParagraph().AppendText("First row, First cell");
                table[0, 1].AddParagraph().AppendText("First row, Second cell");
                table[1, 0].AddParagraph().AppendText("Second row, First cell");
                table[1, 1].AddParagraph().AppendText("Second row, Second cell");
                //Specifies the vertical merge start to first row first cell.
                table[0, 0].CellFormat.VerticalMerge = CellMerge.Start;
                //Modifies the cell content.
                table[0, 0].Paragraphs[0].Text = "Vertically merged cell";
                //Specifies the vertical merge continue to second row first cell.
                table[1, 0].CellFormat.VerticalMerge = CellMerge.Continue;
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
