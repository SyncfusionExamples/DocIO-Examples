using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Data;
using Syncfusion.Drawing;


// Holds row index to color mapping
Dictionary<int, object> TextColors = new Dictionary<int, object>();
// Load the Word document
using (WordDocument document = new WordDocument(Path.GetFullPath(@"Data/Template.docx")))
{

    // Get invoice data
    DataTable invoiceTable = GetInvoiceData();

    // Store color info by row index
    for (int i = 0; i < invoiceTable.Rows.Count; i++)
        TextColors.Add(i, invoiceTable.Rows[i]["FontColor"]);

    // Hook merge event to apply font color
    document.MailMerge.MergeField += ApplyColorToFields;

    // Enable separate page for each invoice
    document.MailMerge.StartAtNewPage = true;

    // Perform mail merge
    document.MailMerge.ExecuteGroup(invoiceTable);

    // Save the modified document
    document.Save(Path.GetFullPath(@"Output/Result.docx"), FormatType.Docx);
}

// Event handler to apply text color based on row
void ApplyColorToFields(object sender, MergeFieldEventArgs args)
{
    if (TextColors.TryGetValue(args.RowIndex, out object color))
        args.TextRange.CharacterFormat.TextColor = (Color)color;
}

// Generates sample invoice data
DataTable GetInvoiceData()
{
    DataTable table = new DataTable("Invoice");

    table.Columns.Add("InvoiceNumber");
    table.Columns.Add("InvoiceDate");
    table.Columns.Add("CustomerName");
    table.Columns.Add("ItemDescription");
    table.Columns.Add("Amount");
    table.Columns.Add("FontColor", typeof(Color));

    // First Invoice
    table.Rows.Add("INV001", "2024-05-01", "Andy Bernard", "Consulting Services", "$3000.00", Color.Teal);
    // Second Invoice
    table.Rows.Add("INV002", "2024-05-05", "Stanley Hudson", "Software Development", "$4500.00", Color.DarkOrange);
    // Third Invoice
    table.Rows.Add("INV003", "2024-05-10", "Margaret Peacock", "UI Design Services", "$2000.00", Color.Indigo);

    return table;
}

