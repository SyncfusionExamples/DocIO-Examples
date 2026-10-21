using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;

namespace Remove_rows_after_specified_row_count_in_table
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an existing Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    //Accesses the instance of the third section in the Word document.
                    WSection section = document.Sections[0];
                    //Accesses the instance of the first table in the section.
                    WTable table = section.Tables[0] as WTable;
                    //Rows as much as data come from database.
                    int numOfRows = 2;
                    //Remove all rows except first three rows including header row (specified number of rows) in the table.
                    for (int i = table.Rows.Count - 1; i > numOfRows; i--)
                    {
                        table.Rows.RemoveAt(i);
                    }
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Output/output.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
