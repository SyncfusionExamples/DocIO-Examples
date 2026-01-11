using Syncfusion.DocIO.DLS;

namespace Cloned_and_add_table_row_with_restart_numbered_list
{
    class Program
    {
        public static void Main(string[] args)
        {
            // Load the existing Word document
            WordDocument document = new WordDocument(Path.GetFullPath(@"Data\Input.docx"));
            // Retrieve the first table from the last section of the document
            WTable table = (WTable)document.LastSection.Tables[0];
            // Clone the third row (index 2) of the table
            WTableRow clonedRow = table.Rows[2].Clone();
            // Insert the cloned row back into the table at position 3 (after the original row)
            table.Rows.Insert(3, clonedRow);
            // Iterate through all cells in the newly inserted row (row index 3)
            foreach (WTableCell cell in table.Rows[3].Cells)
            {
                // Flag to track whether the first list paragraph has been encountered
                bool isListStart = false;
                // Iterate through all paragraphs inside the current cell
                foreach (WParagraph paragraph in cell.Paragraphs)
                {
                    // Check if paragraph is a list
                    if (paragraph.ListFormat.ListType != ListType.NoList)
                    {
                        // If a list has already started, continue numbering to align with the existing list
                        if (isListStart)
                            paragraph.ListFormat.ContinueListNumbering();
                        else
                        {
                            // Mark that the first list paragraph has been found
                            isListStart = true;
                            // Restart numbering for the first list paragraph in the cloned ro
                            paragraph.ListFormat.RestartNumbering = true;
                        }
                    }
                }
            }
            // Save the Word document
            document.Save(Path.GetFullPath("../../../Output/Output.docx"));
            // Close the Word document
            document.Close();
        }
    }
}
