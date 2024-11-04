using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Data;

using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
{
    // Opens the template document.
    using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
    {
        // Uses the mail merge events to skip a range of records during mail merge.
        document.MailMerge.MergeField += new MergeFieldEventHandler(SkipRangeOfRecords);
        // Executes Mail Merge with groups.
        document.MailMerge.ExecuteGroup(GetDataTable());
        using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
        {
            // Saves the Word document to file stream.
            document.Save(outputStream, FormatType.Docx);
        }
    }
}

#region Helper methods
/// <summary>
/// Event handler to skip a range of records during mail merge.
/// </summary>
void SkipRangeOfRecords(object sender, MergeFieldEventArgs args)
{
    // Check if the current row index is between 2 and 7.
    // If so, set the field text to an empty string, effectively skipping those records.
    if (args.RowIndex > 2 && args.RowIndex < 7)
        args.Text = string.Empty;
}

/// <summary>
/// Gets a DataTable populated with employee data for mail merge.
/// </summary>
DataTable GetDataTable()
{
    // Create a new DataTable with columns for employee name and number.
    DataTable dataTable = new DataTable("Employee");
    dataTable.Columns.Add("EmployeeName");
    dataTable.Columns.Add("EmployeeNumber");

    // Populate the DataTable with 20 rows of employee data.
    for (int i = 0; i < 20; i++)
    {
        // Create a new DataRow and add it to the DataTable.
        DataRow datarow = dataTable.NewRow();
        dataTable.Rows.Add(datarow);
        // Set employee name.
        datarow[0] = "Employee" + i.ToString();
        // Set employee number.
        datarow[1] = "EMP" + i.ToString();       
    }
    // Return the populated DataTable.
    return dataTable;  
}

#endregion