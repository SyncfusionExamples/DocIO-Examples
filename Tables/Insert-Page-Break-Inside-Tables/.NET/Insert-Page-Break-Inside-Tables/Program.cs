using Syncfusion.DocIO.DLS;
using System.IO;

namespace Insert_Page_Break_Inside_Tables
{
    class Program
    {
        static void Main(string[] args)
        {
            // Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                // Adds a default section and paragraph to the document.
                document.EnsureMinimal();

                //Adds a new Table to the text body.
                WTable table = document.LastSection.Body.AddTable() as WTable;

                //Insert rows to the table. This will apply the format to a whole table.
                table.ResetCells(5, 5);

                // Sample text collection used to populate table cells.
                string[] randomTexts =
                         {
                            "Lorem ipsum dolor sit amet",
                            "Syncfusion DocIO library",
                            "Word document processing",
                            "Table cell sample text",
                            "Page break demonstration",
                            "Random content generation",
                            "Document automation example",
                            "Text inside table row",
                            "Sample paragraph content",
                            "Testing table layout"
                        };
                // Populate each table cell with sample text.
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    WTableRow row = table.Rows[i];

                    for (int j = 0; j < row.Cells.Count; j++)
                    {
                        // Add a paragraph to the current cell.
                        WParagraph paragraph = row.Cells[j].AddParagraph() as WParagraph;

                        // Add random text
                        paragraph.AppendText(randomTexts[(i * row.Cells.Count + j) % randomTexts.Length]);
                    }
                }
                // Insert page breaks between table rows.
                InserPageBreak(table);
                //saves the document
                document.Save(Path.GetFullPath("Output/Output.docx"));
            }
        }

        /// <summary>
        /// Inserts page break into the tablecell
        /// </summary>
        /// <param name="table"></param>
        private static void InserPageBreak(WTable table)
        {
            int i = 1;
            while (i < table.Rows.Count)
            {
                //To get the owner textbody of the table
                WTextBody body = table.Owner as WTextBody;
                //Adds an empty paragraph and insert the pagebreak
                body.AddParagraph().AppendBreak(Syncfusion.DocIO.DLS.BreakType.PageBreak);
                //Add the new table to the owner textbody
                WTable pageBreaktable = body.AddTable() as WTable;
                //Moves the row to the new table
                WTableRow row = table.Rows[i] as WTableRow;
                pageBreaktable.Rows.Add(row);
            }
        }
    }
}
