using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Set_text_direction_to_table
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Loads the template document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    //Gets the text body of first sectio
                    WTextBody textbody = document.Sections[0].Body;
                    //Gets the table.
                    IWTable table = textbody.Tables[0];
                    //Iterates through table row.
                    foreach (WTableRow row in table.Rows)
                    {
                        foreach (WTableCell cell in row.Cells)
                        {
                            //Sets the text direction for the contents.
                            cell.CellFormat.TextDirection = TextDirection.Vertical;
                        }
                    }
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
