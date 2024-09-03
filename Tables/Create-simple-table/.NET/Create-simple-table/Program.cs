using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Create_simple_table
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Adds a section into Word document.
                IWSection section = document.AddSection();
                //Adds a new paragraph into Word document and appends text into paragraph.
                IWTextRange textRange = section.AddParagraph().AppendText("Price Details");
                textRange.CharacterFormat.FontName = "Arial";
                textRange.CharacterFormat.FontSize = 12;
                textRange.CharacterFormat.Bold = true;
                section.AddParagraph();
                //Adds a new table into Word document.
                IWTable table = section.AddTable();
                //Specifies the total number of rows & columns.
                table.ResetCells(3, 2);
                //Accesses the instance of the cell (first row, first cell) and adds the content into cell.
                textRange = table[0, 0].AddParagraph().AppendText("Item");
                textRange.CharacterFormat.FontName = "Arial";
                textRange.CharacterFormat.FontSize = 12;
                textRange.CharacterFormat.Bold = true;
                //Accesses the instance of the cell (first row, second cell) and adds the content into cell.
                textRange = table[0, 1].AddParagraph().AppendText("Price($)");
                textRange.CharacterFormat.FontName = "Arial";
                textRange.CharacterFormat.FontSize = 12;
                textRange.CharacterFormat.Bold = true;
                //Accesses the instance of the cell (second row, first cell) and adds the content into cell.
                textRange = table[1, 0].AddParagraph().AppendText("Apple");
                textRange.CharacterFormat.FontName = "Arial";
                textRange.CharacterFormat.FontSize = 10;
                //Accesses the instance of the cell (second row, second cell) and adds the content into cell.
                textRange = table[1, 1].AddParagraph().AppendText("50");
                textRange.CharacterFormat.FontName = "Arial";
                textRange.CharacterFormat.FontSize = 10;
                //Accesses the instance of the cell (third row, first cell) and adds the content into cell.
                textRange = table[2, 0].AddParagraph().AppendText("Orange");
                textRange.CharacterFormat.FontName = "Arial";
                textRange.CharacterFormat.FontSize = 10;
                //Accesses the instance of the cell (third row, second cell) and adds the content into cell.
                textRange = table[2, 1].AddParagraph().AppendText("30");
                textRange.CharacterFormat.FontName = "Arial";
                textRange.CharacterFormat.FontSize = 10;
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
