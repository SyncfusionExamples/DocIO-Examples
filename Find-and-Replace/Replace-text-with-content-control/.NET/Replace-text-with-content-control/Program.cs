using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;
using System.Text.RegularExpressions;

namespace Replace_text_with_content_control
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Load an existing Word document.
                using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
                {
                    //Create a BlockContentControl.
                    BlockContentControl blockContentControl = new BlockContentControl(document, ContentControlType.RichText);
                    //Add a new table to the block content control.
                    WTable table = blockContentControl.TextBody.AddTable() as WTable;
                    //Specify the total number of rows and columns.
                    table.ResetCells(1, 2);
                    //Get first row of the table.
                    WTableRow row = table.Rows[0];
                    //Add a new paragraph to the first cell of the table.
                    IWParagraph cellParagraph = row.Cells[0].AddParagraph();
                    row.Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Top;
                    FileStream imageStream = new FileStream(Path.GetFullPath(@"Data/image.png"), FileMode.Open, FileAccess.ReadWrite);
                    //Append a picture to the cell.
                    WPicture picture = cellParagraph.AppendPicture(imageStream) as WPicture;
                    picture.Height = 88.3f;
                    picture.Width = 142.2f;
                    //Add a new paragraph to the next cell.
                    cellParagraph = row.Cells[1].AddParagraph();
                    row.Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Top;
                    cellParagraph.ParagraphFormat.BeforeSpacing = 12f;
                    //Append text to the cell.
                    IWTextRange text = cellParagraph.AppendText("Mountain-200");
                    cellParagraph.ParagraphFormat.AfterSpacing = 3f;
                    //Set the text format.
                    text.CharacterFormat.Bold = true;
                    text.CharacterFormat.FontName = "Arial";
                    text.CharacterFormat.FontSize = 16f;
                    //Add a new paragraph.
                    cellParagraph = row.Cells[1].AddParagraph();
                    //Append text to the paragraph of a cell.
                    cellParagraph.AppendText("Product No: BK-M68B-38\nSize: 38\nWeight: 25\nPrice: $2,294.99\n\n");
                    //Specify the table borders border type.
                    table.TableFormat.Borders.BorderType = BorderStyle.Single;
                    //Create a textbody part object.
                    TextBodyPart textBodyPart = new TextBodyPart(document);
                    //Add the block content control to the textbodyPart.
                    textBodyPart.BodyItems.Add(blockContentControl);
                    //Replace all entries of a given regular expression with the text body part.
                    document.Replace(new Regex("<<(.*)>>"), textBodyPart);
                    //Create a file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to the file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
