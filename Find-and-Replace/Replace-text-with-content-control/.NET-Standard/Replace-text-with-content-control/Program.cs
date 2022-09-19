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
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Data/Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Load an existing Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    //Add block content control into the Word document.
                    BlockContentControl blockContentControl = new BlockContentControl(document, ContentControlType.RichText);
                    //Add a new table to the block content control.
                    WTable table = blockContentControl.TextBody.AddTable() as WTable;
                    //Specify the total number of rows and columns.
                    table.ResetCells(1, 2);
                    //Get first row of the table.
                    WTableRow row = table.Rows[0];
                    //Add a new paragraph to the first cell of the table.
                    IWParagraph cellPara = row.Cells[0].AddParagraph();
                    row.Cells[0].CellFormat.VerticalAlignment = VerticalAlignment.Top;
                    FileStream imageStream = new FileStream(Path.GetFullPath(@"../../../Data/image.png"), FileMode.Open, FileAccess.ReadWrite);
                    //Append a picture to the cell.
                    WPicture picture = cellPara.AppendPicture(imageStream) as WPicture;
                    picture.Height = 88.3f;
                    picture.Width = 142.2f;
                    //Add a new paragraph to the next cell.
                    cellPara = row.Cells[1].AddParagraph();
                    row.Cells[1].CellFormat.VerticalAlignment = VerticalAlignment.Top;
                    cellPara.ParagraphFormat.BeforeSpacing = 12f;
                    //Append text to the cell.
                    IWTextRange txt = cellPara.AppendText("Mountain-200");
                    cellPara.ParagraphFormat.AfterSpacing = 3f;
                    //Set the text format.
                    txt.CharacterFormat.Bold = true;
                    txt.CharacterFormat.FontName = "Arial";
                    txt.CharacterFormat.FontSize = 16f;
                    //Add a new paragraph.
                    cellPara = row.Cells[1].AddParagraph();
                    //Append texts to the paragraph of a cell.
                    txt = cellPara.AppendText("Product No: BK-M68B-38");
                    txt = cellPara.AppendText("\nSize: 38");
                    txt = cellPara.AppendText("\nWeight: 25");
                    txt = cellPara.AppendText("\nPrice: $2,294.99\n\n");
                    //Specify the table borders border type.
                    table.TableFormat.Borders.BorderType = BorderStyle.Single;

                    TextBodyPart textBodyPart = new TextBodyPart(document);
                    textBodyPart.BodyItems.Add(blockContentControl);
                    //Replace all entries of a given regular expression with the text body part.
                    document.Replace(new Regex("<<(.*)>>"), textBodyPart);
                    //Create a file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to the file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
