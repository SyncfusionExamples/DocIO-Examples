using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Licensing;
using System.IO;

namespace Set_table_row_height
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
                    //Gets the text body of first section.
                    WTextBody textbody = document.Sections[0].Body;
                    //Gets the table.
                    IWTable table = textbody.Tables[0];
                    //Iterates through table rows.
                    foreach (WTableRow row in table.Rows)
                    {
                        WTableRow tableRow = row as WTableRow;
                        //Set table row height.
                        tableRow.Height = 30.2f;
                        //Set table row height type.
                        tableRow.HeightType = TableRowHeightType.Exactly;
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
