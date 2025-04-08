using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Licensing;

namespace Keep_paragraphs_together_in_table
{
    internal class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                using (WordDocument document = new WordDocument(fileStream, FormatType.Automatic))
                {
                    //Find all tables by EntityType in Word document.
                    List<Entity> tables = document.FindAllItemsByProperty(EntityType.Table, null, null);

                    for (int i = 0; i < tables.Count; i++)
                    {
                        WTable table = tables[i] as WTable;
                        // Apply "Keep with Next" for all paragraphs inside the table
                        SetKeepTogetherForTable(table);
                    }

                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }

        /// <summary>
        /// Ensures that all paragraphs inside a table stay together.
        /// </summary>
        private static void SetKeepTogetherForTable(WTable table)
        {
            foreach (WTableRow row in table.Rows)
            {
                foreach (WTableCell cell in row.Cells)
                {
                    foreach (Entity item in cell.ChildEntities)
                    {
                        if (item is WTable nestedTable)
                        {
                            // Recursively process nested tables
                            SetKeepTogetherForTable(nestedTable);
                        }
                        else if (item is WParagraph paragraph)
                        {
                            // Ensure paragraphs stay together
                            paragraph.ParagraphFormat.KeepFollow = true; // Keep with next
                        }
                    }
                }
            }
        }
    }
}
