using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Adjust_first_column_width
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Load the input Word document from file stream
            using (FileStream docStream = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open, FileAccess.Read))
            {
                // Open the Word document
                using (WordDocument document = new WordDocument(docStream, FormatType.Docx))
                {
                    // Find all tables in the document
                    List<Entity> tables = document.FindAllItemsByProperty(EntityType.Table, null, null);

                    // Initialize variables
                    bool isFisrtCell = false; // Flag to identify the first cell in each row
                    WTableCell firstCellReference = new WTableCell(document); // To store the reference to the first cell
                    float totalCellWidth = 0; // Accumulate width of cells except the first one

                    // Loop through each table found in the document
                    foreach (WTable table in tables)
                    {
                        // Iterate through each row in the table
                        foreach (WTableRow row in table.Rows)
                        {
                            // Reset variables for each row
                            totalCellWidth = 0;
                            isFisrtCell = false;

                            // Iterate through each cell in the row
                            foreach (WTableCell cell in row.Cells)
                            {
                                if (!isFisrtCell)
                                {
                                    // Identify the first cell in the row
                                    isFisrtCell = true;
                                    firstCellReference = cell; // Store the first cell reference
                                }
                                else
                                {
                                    // Add the width of remaining cells in the row
                                    totalCellWidth += cell.Width;
                                }
                            }

                            // Calculate the remaining width by subtracting totalCellWidth from page width
                            float pageWidth = (table.Owner.Owner as WSection).PageSetup.ClientWidth;
                            float remainingWidth = pageWidth - totalCellWidth;

                            // Set the width of the first cell to the remaining width
                            firstCellReference.Width = remainingWidth;
                        }
                    }

                    // Save the modified document to a new file
                    using (FileStream docStream1 = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.Write))
                    {
                        document.Save(docStream1, FormatType.Docx);
                    }
                }
            }
        }
    }
}
