using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

namespace Add_or_remove_column_in_a_table
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load the Word document
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"Data/Template.docx")))
            {
                // Access the first table in the document
                WTable table = (WTable)document.Sections[0].Tables[0];
                // Add a new column at index
                InsertColumn(table, 1);
                // Add a column at the last index
                AddColumn(table);
                // Remove a column at the index 
                RemoveColumn(table, 4);
                // Save the modified document to a new file
                document.Save(Path.GetFullPath(@"Output/Result.docx"), FormatType.Docx);
            }
        }
        /// <summary>
        /// Adds a new column at the last index in the table
        /// </summary>
        /// <param name="table">The table to modify</param>
        private static void AddColumn(WTable table)
        {
            for (int i = 0; i < table.Rows.Count; i++)
                table.Rows[i].AddCell();
        }
        /// <summary>
        /// Adds a new column at the specified index in the table.
        /// </summary>
        /// <param name="table">The table to modify.</param>
        /// <param name="indexToAdd">The index at which to insert the new column.</param>
        private static void InsertColumn(WTable table, int indexToAdd)
        {
            // Loop through each row in the table
            for (int i = 0; i < table.Rows.Count; i++)
            {
                // Check if the index is within the valid range for the current row
                if (indexToAdd >= 0 && indexToAdd <= table.Rows[i].Cells.Count)
                {
                    // Create a new cell.
                    WTableCell newCell = new WTableCell(table.Document);
                    // Insert the new cell at the specified index in the current row
                    table.Rows[i].Cells.Insert(indexToAdd, newCell);
                }                
            }
        }
        /// <summary>
        /// Removes a column at the specified index from the table.
        /// </summary>
        /// <param name="table">The table to modify.</param>
        /// <param name="indexToRemove">The index of the column to remove.</param>
        private static void RemoveColumn(WTable table, int indexToRemove)
        {
            // Loop through each row in the table
            for (int i = 0; i < table.Rows.Count; i++)
            {
                // Check if the index is within the valid range for the current row
                if (indexToRemove >= 0 && indexToRemove < table.Rows[i].Cells.Count)
                {
                    // Remove the cell at the specified index in the current row
                    table.Rows[i].Cells.Remove(table.Rows[i].Cells[indexToRemove]);
                }
            }
        }
    }
}
