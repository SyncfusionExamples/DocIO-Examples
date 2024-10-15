using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Dynamic;

using (FileStream inputFileStream = new FileStream(Path.GetFullPath("Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
{
    // Open the input Word document
    using (WordDocument document = new WordDocument(inputFileStream, FormatType.Docx))
    {
        // Find a table with its alternate text (Title property).
        WTable table = document.FindItemByProperty(EntityType.Table, "Title", "DataTable") as WTable;
        // Check if the table was found.
        if (table != null)
        {
            // Get the second row of the table.
            WTableRow secondRow = table.Rows[1];
            // Insert data into the cells of the second row.
            InsertDataToCells(secondRow);
            // Add dynamic rows starting at index 2, based on the second row.
            AddDyamicRows(table, 2, secondRow);
        }

        using (FileStream outputFileStream = new FileStream(Path.GetFullPath("Output/Result.docx"), FileMode.Create, FileAccess.Write))
        {
            // Save the modified document to the output file stream.
            document.Save(outputFileStream, FormatType.Docx);
        }
    }
}

/// <summary>
/// Insert data into the cells of a specified table row.
/// </summary>
void InsertDataToCells(WTableRow row)
{
    // List of placeholder data to insert into the cells.
    List<string> data = new List<string> { "<<Data1>>", "<<Data2>>", "<<Data3>>", "<<Data4>>" };
    int count = 0;
    // Iterate through each cell in the specified row.
    foreach (WTableCell cell in row.Cells)
    {
        // Assign data to the particular cell.
        cell.Paragraphs[0].Text = data[count];
        count++;
    }
}

/// <summary>
/// Add dynamic rows to a specified table at a certain index.
/// </summary>
void AddDyamicRows(WTable table, int index, WTableRow row)
{
    // Create a list of dynamic row details.
    IEnumerable<dynamic> rowsDetails = CreateDyamicRows();
    // Iterate through each dynamic row detail.
    foreach (dynamic rowDetails in rowsDetails)
    {
        // Retrieve cell content for the new row.
        List<string> cellDetails = GetListOfCellValue(rowDetails);
        // Clone the second row to create a new row.
        WTableRow newRow = row.Clone();
        // Iterate through the cells of the cloned row.
        for (int i = 0; i < newRow.Cells.Count; i++)
        {
            // Get the cell at specific from the cloned row.
            WTableCell wTableCell = newRow.Cells[i];
            // Modify the paragraph text of the cell with the corresponding cell detail.
            wTableCell.Paragraphs[0].Text = cellDetails[i];
        }
        // Insert the newly created row at the specified index.
        table.Rows.Insert(index, newRow);
        // Increment the index for the next dynamic row.
        index++;
    }
}

/// <summary>
/// Create dynamic rows with sample cell data.
/// </summary>
IEnumerable<dynamic> CreateDyamicRows()
{
    // Create a list of dynamic row details.
    List<dynamic> rowDetails = new List<dynamic>();

    // Add dynamic cells to the row details list.
    rowDetails.Add(CreateDynamicCells("<<Data5>>", "<<Data6>>", "<<Data7>>", "<<Data8>>"));
    rowDetails.Add(CreateDynamicCells("<<Data9>>", "<<Data10>>", "<<Data11>>", "<<Data12>>"));
    rowDetails.Add(CreateDynamicCells("<<Data13>>", "<<Data14>>", "<<Data15>>", "<<Data16>>"));
    rowDetails.Add(CreateDynamicCells("<<Data17>>", "<<Data18>>", "<<Data19>>", "<<Data20>>"));
    // Return the list of dynamic row details.
    return rowDetails; 
}

/// <summary>
/// Create dynamic cell data.
/// </summary>
dynamic CreateDynamicCells(string cell1, string cell2, string cell3, string cell4)
{
    // Create a new ExpandoObject for dynamic properties.
    dynamic dynamicOrder = new ExpandoObject(); 

    // Assign values to the dynamic object properties.
    dynamicOrder.Cell1 = cell1;
    dynamicOrder.Cell2 = cell2;
    dynamicOrder.Cell3 = cell3;
    dynamicOrder.Cell4 = cell4;
    // Return the dynamic object.
    return dynamicOrder;
}

/// <summary>
/// Convert the dynamic values to a list of strings.
/// </summary>
List<string> GetListOfCellValue(dynamic rowDetails)
{
    List<string> cellDetails = new List<string>();

    // Add each dynamic cell value to the list.
    cellDetails.Add(rowDetails.Cell1);
    cellDetails.Add(rowDetails.Cell2);
    cellDetails.Add(rowDetails.Cell3);
    cellDetails.Add(rowDetails.Cell4);
    // Return the list of cell details.
    return cellDetails;
}
