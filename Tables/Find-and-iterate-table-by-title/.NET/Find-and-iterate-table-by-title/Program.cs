using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;

namespace Find_and_iterate_table_by_title
{
    internal class Program
    {
        static void Main(string[] args)
        {        
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Opens an existing Word document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Automatic))
                {
                    // Find the table with title.
                    WTable table = document.FindItemByProperty(EntityType.Table, "Title", "Overview") as WTable;
                    if (table != null)
                    {
                        // Iterate through the rows and cells of the table
                        foreach (WTableRow row in table.Rows)
                        {
                            //Iterates through the cells of rows.
                            foreach (WTableCell cell in row.Cells)
                            {
                                //Iterates through the paragraphs of the cell.
                                foreach (WParagraph paragraph in cell.Paragraphs)
                                {
                                    //When the paragraph contains text Panda then insert new text into paragraph.
                                    if (paragraph.Text.Contains("panda"))
                                    {
                                        WTextRange insertedText = paragraph.AppendText(" (Attributes)") as WTextRange;
                                        // Apply simple formatting only to the inserted text
                                        insertedText.CharacterFormat.Bold = true;
                                        insertedText.CharacterFormat.TextColor = Color.Red;
                                    }
                                }
                            }
                        }
                    }
                    //Creates file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}

