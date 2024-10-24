using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

using (FileStream inputStream = new FileStream("Data/Template.docx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    // Open the template Word document.
    using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
    {
        // Access the body of the first section of the document.
        WTextBody body = document.Sections[0].Body;
        // Retrieve the first table from the body of the document.
        WTable table = document.Sections[0].Tables[0] as WTable;
        // Convert the table to text.
        string convertedText = ConvertTableToText(document, table);
        // Split the converted text by line breaks to create individual paragraphs.
        string[] paraText = convertedText.Split(new string[] { "\r\n" }, StringSplitOptions.None);
        // Get the index of the table within the body entities.
        int index = body.ChildEntities.IndexOf(table);
        // Remove the table from the document body.
        body.ChildEntities.Remove(table);
        // Insert each line of the converted text as a new paragraph at the table's original location.
        for (int i = 0; i < paraText.Length; i++)
        {
            // Create a new paragraph.
            IWParagraph paragraph = new WParagraph(document);
            // Append the corresponding line of text to the paragraph.
            paragraph.AppendText(paraText[i]);
            // Insert the paragraph into the body at the correct index.
            body.ChildEntities.Insert(index + i, paragraph);
        }
        // Save the modified document.
        using (FileStream outputStream = new FileStream("Output/Output.docx", FileMode.Create, FileAccess.Write))
        {
            document.Save(outputStream, FormatType.Docx);
        }
    }
}

/// <summary>
/// Converts the specified table into text format.
/// </summary>
static string ConvertTableToText(WordDocument document, WTable table)
{
    // Use StringWriter to build the text output.
    StringWriter stringWriter = new StringWriter();
    // Iterate through each row of the table.
    for (int row = 0; row < table.Rows.Count; row++)
    {
        // Iterate through each cell in the current row.
        for (int column = 0; column < table.Rows[row].Cells.Count; column++)
        {
            // Get the current cell.
            WTableCell cell = table.Rows[row].Cells[column];
            // Iterate through each paragraph in the cell.
            for (int para = 0; para < cell.Paragraphs.Count; para++)
            {
                // Get the current paragraph.
                IWParagraph paragraph = cell.Paragraphs[para] as IWParagraph;
                // Iterate through each item in the paragraph.
                for (int item = 0; item < paragraph.Items.Count; item++)
                {
                    // If the item is a text range, write the text to the StringWriter.
                    if (paragraph.Items[item] is IWTextRange textRange)
                    {
                        stringWriter.Write(textRange.Text);
                    }
                }
                // Add a new line for each paragraph within the same cell, except the last one.
                if (para < cell.Paragraphs.Count - 1)
                {
                    stringWriter.Write("\r\n");
                }
            }
            // Add a comma delimiter between cells, except for the last cell in the row.
            if (column < table.Rows[row].Cells.Count - 1)
            {
                stringWriter.Write(",");
            }
        }
        // Add a new line for each row, except for the last row.
        if (row < table.Rows.Count - 1)
        {
            stringWriter.Write("\r\n");
        }
    }
    // Return the built string representing the table content.
    return stringWriter.ToString();
}
