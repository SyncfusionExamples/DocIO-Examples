using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Dynamic_table_by_adding_rows
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
                section.AddParagraph();
                //Adds a new table into Word document.
                IWTable table = section.AddTable();
                //Adds the first row into table.
                WTableRow row = table.AddRow();
                //Adds the first cell into first row.
                WTableCell cell = row.AddCell();
                //Specifies the cell width.
                cell.Width = 200;
                cell.AddParagraph().AppendText("Item");
                //Adds the second cell into first row.
                cell = row.AddCell();
                //Specifies the cell width.
                cell.Width = 200;
                cell.AddParagraph().AppendText("Price($)");
                //Adds the second row into table.
                row = table.AddRow(true, false);
                //Adds the first cell into second row.
                cell = row.AddCell();
                //Specifies the cell width.
                cell.Width = 200;
                cell.AddParagraph().AppendText("Apple");
                //Adds the second cell into second row.
                cell = row.AddCell();
                //Specifies the cell width.
                cell.Width = 200;
                cell.AddParagraph().AppendText("50");
                //Adds the third row into table.
                row = table.AddRow(true, false);
                //Adds the first cell into third row.
                cell = row.AddCell();
                //Specifies the cell width.
                cell.Width = 200;
                cell.AddParagraph().AppendText("Orange");
                //Adds the second cell into third row.
                cell = row.AddCell();
                //Specifies the cell width.
                cell.Width = 200;
                cell.AddParagraph().AppendText("30");
                //Adds the fourth row into table.
                row = table.AddRow(true, false);
                //Adds the first cell into fourth row.
                cell = row.AddCell();
                //Specifies the cell width.
                cell.Width = 200;
                cell.AddParagraph().AppendText("Banana");
                //Adds the second cell into fourth row.
                cell = row.AddCell();
                //Specifies the cell width.
                cell.Width = 200;
                cell.AddParagraph().AppendText("20");
                //Adds the fifth row to table.
                row = table.AddRow(true, false);
                //Adds the first cell into fifth row.
                cell = row.AddCell();
                //Specifies the cell width.
                cell.Width = 200;
                cell.AddParagraph().AppendText("Grapes");
                //Adds the second cell into fifth row.
                cell = row.AddCell();
                //Specifies the cell width.
                cell.Width = 200;
                cell.AddParagraph().AppendText("70");
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
