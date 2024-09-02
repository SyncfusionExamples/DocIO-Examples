using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Set_table_cell_width
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
                        //Sets width for cells.
                        for (int i = 0; i < row.Cells.Count; i++)
                        {
                            WTableCell cell = row.Cells[i];
                            if (i % 2 == 0)
                                //Sets width as 100 for cells in even column.
                                cell.Width = 100;
                            else
                                //Sets width as 150 for cell in odd column.
                                cell.Width = 150;
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
