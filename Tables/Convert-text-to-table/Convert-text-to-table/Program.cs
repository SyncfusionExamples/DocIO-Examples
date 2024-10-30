using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using Syncfusion.Drawing;

using (FileStream inputStream = new FileStream("Data/Input.docx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    // Open the input Word document.
    using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
    {
        // Access the body of the first section.
        WTextBody body = document.Sections[0].Body;
        // Retrieve the paragraphs that is need to be converted into a table.
        WParagraph para1 = body.Paragraphs[0];
        WParagraph para2 = body.Paragraphs[1];
        WParagraph para3 = body.Paragraphs[2];
        WParagraph para4 = body.Paragraphs[3];
        // Combine the text from the paragraphs, each separated by a new line.
        string text = para1.Text + "\r\n" + para2.Text + "\r\n" + para3.Text + "\r\n" + para4.Text;
        // Convert the combined text into a table.
        IWTable table = ConvertTextToTable(document, text);
        // Get the index of the first paragraph to insert the table at the correct position.
        int index = body.ChildEntities.IndexOf(para1);
        // Remove the selected paragraphs from the document.
        body.ChildEntities.Remove(para1);
        body.ChildEntities.Remove(para2);
        body.ChildEntities.Remove(para3);
        body.ChildEntities.Remove(para4);
        // Insert the newly created table at the location of the first paragraph.
        body.ChildEntities.Insert(index, table);
        // Save the modified document with the table inserted.
        using (FileStream outputStream = new FileStream("Output/Output.docx", FileMode.Create, FileAccess.Write))
        {
            document.Save(outputStream, FormatType.Docx);
        }
    }
}

/// <summary>
/// Converts the provided text into a table format with rows and columns.
/// </summary>
IWTable ConvertTextToTable(WordDocument document, string text)
{
    // Split the text into rows based on line breaks.
    string[] rows = text.Split(new string[] { "\r\n" }, StringSplitOptions.None);
    // Create a new table in the document.
    IWTable table = new WTable(document);
    // Determine the number of rows and columns based on the text structure.
    int rowCount = rows.Length;
    int colCount = rows[0].Split(new char[] { ',' }).Length;
    // Initialize the table with the appropriate number of rows and columns.
    table.ResetCells(rowCount, colCount);
    // Populate the table with the text by iterating through each row and column.
    for (int i = 0; i < rowCount; i++)
    {
        // Split the current row into columns using a comma as the delimiter.
        string[] columns = rows[i].Split(new char[] { ',' });
        for (int j = 0; j < colCount; j++)
        {
            // Add a paragraph to the cell and append the text.
            table[i, j].AddParagraph().AppendText(columns[j].Trim());
        }
    }
    // Set the table's border color to black.
    table.TableFormat.Borders.Color = Color.Black;
    // Set the table's border style to single line.
    table.TableFormat.Borders.BorderType = BorderStyle.Single;
    // Return the created table.
    return table;
}
