using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Fix_image_position
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Open the Word document
            using (WordDocument document = new WordDocument())
            {
                //Adds a section into Word document
                IWSection section = document.AddSection();
                //Adds a new table into Word document
                IWTable table = section.AddTable();
                //Specifies the total number of rows & columns
                table.ResetCells(2, 2);
                //Remove borders
                table.TableFormat.Borders.BorderType = BorderStyle.None;

                //Iterate through all rows
                foreach (WTableRow row in table.Rows)
                {
                    //Iterate through all cells
                    foreach (WTableCell cell in row.Cells)
                    {
                        //Add Image to the cell
                        AddImageToCell(cell);
                    }
                }
                // Save the modified document to a new file
                using (FileStream docStream1 = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.Write))
                {
                    document.Save(docStream1, FormatType.Docx);
                }
            }
        }
        ///<summary>
        ///Add image to a table cell
        ///</summary>
        static void AddImageToCell(WTableCell cell)
        {
            //Specifies the same padding as table option as false to preserve current cell padding
            cell.CellFormat.SamePaddingsAsTable = false;
            //Add cell paddings
            cell.CellFormat.Paddings.Left = 10f;
            cell.CellFormat.Paddings.Right = 10f;
            cell.CellFormat.Paddings.Top = 20f;
            cell.CellFormat.Paddings.Bottom = 20f;

            //Adds the image into the cell.
            FileStream imageStream = new FileStream(Path.GetFullPath(@"Data/Image.png"), FileMode.Open, FileAccess.ReadWrite);
            IWPicture picture = cell.AddParagraph().AppendPicture(imageStream);
            //Set height and width for the image
            picture.Height = 220;
            picture.Width = 220;
            //Close the image stream
            imageStream.Close();
        }
    }
}
