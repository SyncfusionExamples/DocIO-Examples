using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

using (FileStream inputStream = new FileStream(@"Data/Template.docx", FileMode.Open, FileAccess.Read))
{
    // Open an existing Word document.
    using (WordDocument document = new WordDocument(inputStream, FormatType.Automatic))
    {
        // Find the table with title.
        WTable table = document.FindItemByProperty(EntityType.Table, "Title", "Template") as WTable;
        // Check if the table with the specified title was found.
        if (table != null)
        {
            // Define new data to add to the table.
            List<string[]> rowData =
            [
                ["3.", "Banana", "20"],
                ["4.", "Grapes", "70"]
            ];
            for (int j = 0; j < rowData.Count; j++)
            {
                string[] data = rowData[j];
                // Add a new row to the table.
                WTableRow row = table.AddRow(false);
                for (int i = 0; i < data.Length; i++)
                {
                    string cellData = data[i];
                    // Get the cell at the current index.
                    WTableCell cell = row.Cells[i];
                    // Add a paragraph to the cell.
                    WParagraph cellParagraph = cell.AddParagraph() as WParagraph;
                    // Add the cell data as text within the cell's paragraph.
                    cellParagraph.AppendText(cellData);
                }
            }
        }
        using (FileStream outputStream = new FileStream(@"Output/Result.docx", FileMode.Create, FileAccess.Write))
        {
            // Save the modified Word document.
            document.Save(outputStream, FormatType.Docx);
        }
    }
}
