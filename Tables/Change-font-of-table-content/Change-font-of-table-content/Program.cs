using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

using (FileStream inputStream = new FileStream("Data/Input.docx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    // Open the input HTML format document.
    using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
    {
        // Find a table by Title.
        WTable table = document.FindItemByProperty(EntityType.Table, "Title", "Adventure") as WTable;
        // Iterate through each row in the table.
        foreach (WTableRow row in table.Rows)
        {
            // Iterate through each cell in the row.
            foreach (WTableCell cell in row.Cells)
            {
                // Iterate through each paragraph in the cell.
                foreach (WParagraph paragraph in cell.Paragraphs)
                {
                    // Iterate through the child entities of the paragraph.
                    foreach (Entity entity in paragraph.ChildEntities)
                    {
                        // Check if the child entity is a text range.
                        if (entity is WTextRange)
                        {
                            // Apply character format to change the font to Algerian for the text range.
                            (entity as WTextRange).CharacterFormat.FontName = "Algerian";
                        }
                    }
                }
            }
        }
        // Save the modified Word document.
        using (FileStream outputStream = new FileStream("Output/Output.docx", FileMode.Create, FileAccess.Write))
        {
            document.Save(outputStream, FormatType.Docx);
        }
    }
}
