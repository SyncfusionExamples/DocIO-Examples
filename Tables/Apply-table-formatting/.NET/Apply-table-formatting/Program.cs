using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;

namespace Apply_table_formatting
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Creates an instance of WordDocument class (Empty Word Document).
                using (WordDocument document = new WordDocument())
                {
                    //Opens an existing Word document into DocIO instance.
                    document.Open(fileStreamPath, FormatType.Docx);
                    //Accesses the instance of the first section in the Word document.
                    WSection section = document.Sections[0];
                    //Accesses the instance of the first table in the section.
                    WTable table = section.Tables[0] as WTable;
                    //Specifies the title for the table.
                    table.Title = "PriceDetails";
                    //Specifies the description of the table.
                    table.Description = "This table shows the price details of various fruits";
                    //Specifies the left indent of the table.
                    table.IndentFromLeft = 50;
                    //Specifies the background color of the table.
                    table.TableFormat.BackColor = Color.FromArgb(192, 192, 192);
                    //Specifies the horizontal alignment of the table.
                    table.TableFormat.HorizontalAlignment = RowAlignment.Left;
                    //Specifies the left, right, top and bottom padding of all the cells in the table.
                    table.TableFormat.Paddings.All = 10;
                    //Specifies the auto resize of table to automatically resize all cell width based on its content.
                    table.TableFormat.IsAutoResized = true;
                    //Specifies the table top, bottom, left and right border line width.
                    table.TableFormat.Borders.LineWidth = 2f;
                    //Specifies the table horizontal border line width.
                    table.TableFormat.Borders.Horizontal.LineWidth = 2f;
                    //Specifies the table vertical border line width.
                    table.TableFormat.Borders.Vertical.LineWidth = 2f;
                    //Specifies the tables top, bottom, left and right border color.
                    table.TableFormat.Borders.Color = Color.Red;
                    //Specifies the table Horizontal border color.
                    table.TableFormat.Borders.Horizontal.Color = Color.Red;
                    //Specifies the table vertical border color.
                    table.TableFormat.Borders.Vertical.Color = Color.Red;
                    //Specifies the table borders border type.
                    table.TableFormat.Borders.BorderType = BorderStyle.Double;
                    //Accesses the instance of the first row in the table.
                    WTableRow row = table.Rows[0];
                    //Specifies the row height.
                    row.Height = 20;
                    //Specifies the row height type.
                    row.HeightType = TableRowHeightType.AtLeast;
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
