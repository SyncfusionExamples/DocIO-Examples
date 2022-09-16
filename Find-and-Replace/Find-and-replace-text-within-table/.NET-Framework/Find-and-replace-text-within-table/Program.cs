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
            Program program = new Program();  
            //Load the Input document.
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"../../Input.docx"), FormatType.Docx))
            {              
                //Access table in Word document.
                WTable table = document.Sections[0].Tables[0] as WTable;
                program.FindTextWithinTable(table, document); 
                //Save the document.
                document.Save(Path.GetFullPath("../../Sample.docx"), FormatType.Docx);
            }
        }
        /// <summary>
        /// Find text within table in Word document.
        /// </summary>
        /// <param name="table"></param>
        /// <param name="paragraph"></param>
        public void FindTextWithinTable(WTable table, WordDocument document)
        {
            //Iterate through the rows of table.
            foreach (WTableRow row in table.Rows)
            {
                //Iterate through the cells of rows.
                foreach (WTableCell cell in row.Cells)
                {
                    //Iterates through the childentities of the cell as paragraph or table.
                    foreach (Entity entity in cell.ChildEntities)
                    {
                        if (entity.EntityType == EntityType.Paragraph)
                        {
                            WParagraph wParagraph=entity as WParagraph;
                            //Find all the occurrences of a particular text
                            TextSelection[] textSelections = document.FindAll("Suppliers", false, true);
                            for (int i = 0; i < textSelections.Length; i++)
                            {
                                //Replace the specified regular expression with a TextSelection in the paragraph.
                                wParagraph.Replace(new Regex("^//(.*)"), textSelections[i]);
                            }
                        }
                        else if (entity.EntityType == EntityType.Table)
                        {
                            FindTextWithinTable(entity as WTable, document);
                        }
                    }
                }
            }
        }                     
    }
}
