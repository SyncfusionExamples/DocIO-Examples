using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;
using System.Text.RegularExpressions;

namespace Find_and_replace_text_within_table
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the file as Stream.
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Data/Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Load the Input document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    //Access table in Word document.
                    WTable table = document.Sections[0].Tables[0] as WTable;
                    FindAndReplaceInTable(table, document);
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath("../../../Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the document.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }

        }
        /// <summary>
        /// Find and replace text within table in the Word document.
        /// </summary>
        /// <param name="table"></param>
        /// <param name="paragraph"></param>
        private static void FindAndReplaceInTable(WTable table, WordDocument document)
        {
            //Iterate through the rows of table.
            foreach (WTableRow row in table.Rows)
            {
                //Iterate through the cells of rows.
                foreach (WTableCell cell in row.Cells)
                {
                    //Iterates through the Childentities of the cell as paragraph or table.
                    foreach (Entity entity in cell.ChildEntities)
                    {
                        if (entity.EntityType == EntityType.Paragraph)
                        {
                            WParagraph wParagraph = entity as WParagraph;
                            //Find the first occurrence of a particular text in the Word document.
                            TextSelection textSelection = document.Find("Suppliers", false, true);
                            //Replace the specified regular expression with a TextSelection in the paragraph.
                            wParagraph.Replace(new Regex("^//(.*)"), textSelection);
                            
                        }
                        else if (entity.EntityType == EntityType.Table)
                        {
                            FindAndReplaceInTable(entity as WTable, document);
                        }
                    }
                }
            }
        }
    }
}
