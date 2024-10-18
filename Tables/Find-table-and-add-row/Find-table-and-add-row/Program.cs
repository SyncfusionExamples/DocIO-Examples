using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

using (FileStream inputFileStream = new FileStream(Path.GetFullPath("Data/Input.docx"), FileMode.Open, FileAccess.ReadWrite))
{
    // Open the input Word document.
    using (WordDocument document = new WordDocument(inputFileStream, FormatType.Docx))
    {
        //Find a table by Title.
        WTable table = document.FindItemByProperty(EntityType.Table, "Title", "DataTable") as WTable;
        // Check if the table exists.
        if (table != null)
        {
            // Add a new row to the table.
            table.AddRow();
        }
        using (FileStream outputFileStream = new FileStream(Path.GetFullPath("Output/Result.docx"), FileMode.Create, FileAccess.Write))
        {
            // Save the modified document to the output file stream.
            document.Save(outputFileStream, FormatType.Docx);
        }
    }
}