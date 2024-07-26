using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;

namespace Iterating_through_table_elements
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an existing Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    WSection section = document.Sections[0];
                    WTable table = section.Tables[0] as WTable;
                    //Iterates the rows of the table.
                    foreach (WTableRow row in table.Rows)
                    {
                        //Iterates through the cells of rows.
                        foreach (WTableCell cell in row.Cells)
                        {
                            //Iterates through the paragraphs of the cell.
                            foreach (WParagraph paragraph in cell.Paragraphs)
                            {
                                //When the paragraph contains text Panda then apply green as back color to cell.
                                if (paragraph.Text.Contains("panda"))
                                    cell.CellFormat.BackColor = Color.Green;
                            }
                        }
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
}
