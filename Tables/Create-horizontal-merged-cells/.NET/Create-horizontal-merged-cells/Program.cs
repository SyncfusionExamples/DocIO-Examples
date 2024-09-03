using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Create_horizontal_merged_cells
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                IWSection section = document.AddSection();
                section.AddParagraph().AppendText("Horizontal merging of Table cells");
                IWTable table = section.AddTable();
                table.ResetCells(2, 2);
                //Adds content to table cell.
                table[0, 0].AddParagraph().AppendText("First row, First cell");
                table[0, 1].AddParagraph().AppendText("First row, Second cell");
                table[1, 0].AddParagraph().AppendText("Second row, First cell");
                table[1, 1].AddParagraph().AppendText("Second row, Second cell");
                //Specifies the horizontal merge start to first row, first cell.
                table[0, 0].CellFormat.HorizontalMerge = CellMerge.Start;
                //Modifies the cell content.
                table[0, 0].Paragraphs[0].Text = "Horizontally merged cell";
                //Specifies the horizontal merge continue to second row second cell.
                table[0, 1].CellFormat.HorizontalMerge = CellMerge.Continue;
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
