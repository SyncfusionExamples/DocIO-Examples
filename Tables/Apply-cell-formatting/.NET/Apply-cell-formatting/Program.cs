using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;

namespace Apply_cell_formatting
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an existing Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    WSection section = document.Sections[0];
                    WTable table = section.Tables[0] as WTable;
                    //Accesses the instance of the first row in the table.
                    WTableRow row = table.Rows[0];
                    //Specifies the row height.
                    row.Height = 30;
                    //Specifies the row height type.
                    row.HeightType = TableRowHeightType.AtLeast;
                    //Accesses the instance of the first cell in the row.
                    WTableCell cell = row.Cells[0];
                    //Specifies the cell back ground color.
                    cell.CellFormat.BackColor = Color.FromArgb(192, 192, 192);
                    //Specifies the same padding as table option as false to preserve current cell padding.
                    cell.CellFormat.SamePaddingsAsTable = false;
                    //Specifies the left, right, top and bottom padding of the cell.
                    cell.CellFormat.Paddings.Left = 5;
                    cell.CellFormat.Paddings.Right = 5;
                    cell.CellFormat.Paddings.Top = 5;
                    cell.CellFormat.Paddings.Bottom = 5;
                    //Specifies the vertical alignment of content of text.
                    cell.CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    //Disables the text wrap option to avoid displaying longer text on multiple lines.
                    cell.CellFormat.TextWrap = false;
                    // Set text direction for each cell in a row
                    row.Cells[0].CellFormat.TextDirection = TextDirection.Vertical;
                    row.Cells[1].CellFormat.TextDirection = TextDirection.VerticalBottomToTop;
                    row.Cells[2].CellFormat.TextDirection = TextDirection.VerticalTopToBottom;
                    row.Cells[3].CellFormat.TextDirection = TextDirection.VerticalFarEast;
                    row.Cells[4].CellFormat.TextDirection = TextDirection.HorizontalFarEast;
                    row.Cells[5].CellFormat.TextDirection = TextDirection.Horizontal;
                    //Accesses the instance of the second cell in the row.
                    cell = row.Cells[1];
                    cell.CellFormat.BackColor = Color.FromArgb(192, 192, 192);
                    cell.CellFormat.SamePaddingsAsTable = false;
                    //Specifies the left, right, top and bottom padding of the cell.
                    cell.CellFormat.Paddings.All = 5;
                    cell.CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    //Disables the text wrap option to avoid displaying longer text on multiple lines.
                    cell.CellFormat.TextWrap = false;
                    //Access the instance of the third cell in the row.
                    cell = row.Cells[2];
                    //Set color for tablecell borders.
                    cell.CellFormat.Borders.BorderType = BorderStyle.Thick;
                    cell.CellFormat.Borders.Color = Color.Red;
                    cell.CellFormat.Borders.Top.Color = Color.Red;
                    cell.CellFormat.Borders.Bottom.Color = Color.Red;
                    cell.CellFormat.Borders.Right.Color = Color.Red;
                    cell.CellFormat.Borders.Left.Color = Color.Red;
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
}
