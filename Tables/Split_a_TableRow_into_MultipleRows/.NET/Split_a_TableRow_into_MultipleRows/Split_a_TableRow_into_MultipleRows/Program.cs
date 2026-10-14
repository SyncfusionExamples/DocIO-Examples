using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

//Open an existing Word document.
using (FileStream fileStreamPath = new FileStream(@"../../../Data/Input.docx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    // Load the file stream
    using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
    {
        //Get the table from a Word document.
        WTable table = document.Sections[0].Tables[0] as WTable;
        //Split the row into required number of rows
        ConvertOneRowIntoMultipleRows(table, 5, 3);
        //Save a Word document.
        using (FileStream outputStream = new FileStream(@"../../../Output/Result.docx", FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
        {
            document.Save(outputStream, FormatType.Docx);
        }
    }
}


static void ConvertOneRowIntoMultipleRows(WTable table, int rowIndex, int numberOfLines)
{
    for (int i = numberOfLines; i > 0; i--)
    {
        //Clone the row.
        WTableRow row = table.Rows[rowIndex - 1].Clone();
        //Iterate all cells in a row and clear the contents.
        for (int j = 0; j < row.Cells.Count; j++)
        {
            WTableCell tableCell = row.Cells[j];
            tableCell.ChildEntities.Clear();
        }
        //Insert the cloned row in the required index
        table.Rows.Insert(rowIndex, row);
    }
}