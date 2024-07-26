using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Create_table_with_header_row
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                IWSection section = document.AddSection();
                IWTable table = section.AddTable();
                table.ResetCells(50, 1);
                WTableRow row = table.Rows[0];
                //Specifies the first row as a header row of the table.
                row.IsHeader = true;
                row.Height = 20;
                row.HeightType = TableRowHeightType.AtLeast;
                row.Cells[0].AddParagraph().AppendText("Header Row");
                for (int i = 1; i < 50; i++)
                {
                    row = table.Rows[i];
                    row.Height = 20;
                    row.HeightType = TableRowHeightType.AtLeast;
                    row.Cells[0].AddParagraph().AppendText("Text in Row" + i.ToString());
                }
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
