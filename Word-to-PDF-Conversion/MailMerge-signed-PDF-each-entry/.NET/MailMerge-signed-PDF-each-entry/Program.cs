using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Interactive;
using Syncfusion.Pdf.Parsing;
using Syncfusion.Drawing;


// Load the Word document
using (WordDocument document = new WordDocument(Path.GetFullPath(@"Data/Template.docx")))
{
    // Get the recipient details as a list of .NET objects
    List<Recipient> recipients = GetRecipients();

    // Perform mail merge using the recipient data
    document.MailMerge.Execute(recipients);

    // Initialize the DocIO to PDF converter
    using (DocIORenderer renderer = new DocIORenderer())
    {
        // Preserve form fields (textboxes, checkboxes, etc.)
        renderer.Settings.PreserveFormFields = true;

        // Convert the merged Word document to PDF
        using (PdfDocument pdfDocument = renderer.ConvertToPDF(document))
        {
            // Save the PDF to memory stream
            using (MemoryStream stream = new MemoryStream())
            {
                pdfDocument.Save(stream);
                stream.Position = 0;

                // Add digital signature fields in place of textboxes
                AddSignatureField(stream, "Signature");
            }
        }
    }
}


#region Mail Merge Data
/// <summary>
/// Gets the data to perform mail merge.
/// </summary>
/// <returns>List of recipient details.</returns>
List<Recipient> GetRecipients()
{
    List<Recipient> recipients = new List<Recipient>();
    // Initialize sample recipient data
    recipients.Add(new Recipient("Nancy Davolio", "NorthWinds", "507 - 20th Ave. E.Apt. 2A", "507-345-2309"));
    recipients.Add(new Recipient("Janet Leverling", "NorthWinds", "722 Moss Bay Blvd.", "542-754-2843"));
    return recipients;
}
#endregion

#region Add Digital Signature Field
/// <summary>
/// Replaces text box fields in the PDF with signature fields at the same positions.
/// </summary>
/// <param name="stream">PDF stream to modify.</param>
/// <param name="bookmarkTitle">Placeholder text to identify the field (not used directly in this sample).</param>
void AddSignatureField(MemoryStream stream,  string bookmarkTitle)
{
    // Load the generated PDF document
    using (PdfLoadedDocument loadedDocument = new PdfLoadedDocument(stream))
    {
        // Access the form fields in the PDF
        PdfLoadedForm loadedForm = loadedDocument.Form;

        // Loop through each form field
        for (int i = loadedForm.Fields.Count - 1; i >= 0; i--)
        {
            if (loadedForm.Fields[i] is PdfLoadedTextBoxField)
            {
                // Get the loaded text box field
                PdfLoadedTextBoxField loadedTextBoxField = loadedForm.Fields[i] as PdfLoadedTextBoxField;

                // Get its bounds and page to use for placing the signature field
                RectangleF bounds = loadedTextBoxField.Bounds;
                PdfPageBase loadedPage = loadedTextBoxField.Page;

                // Create a signature field with the same name and location
                PdfSignatureField signatureField = new PdfSignatureField(loadedPage, loadedTextBoxField.Text.Trim());
                signatureField.Bounds = bounds;

                // Add the signature field to the document
                loadedDocument.Form.Fields.Add(signatureField);

                // Remove the original textbox field
                loadedDocument.Form.Fields.Remove(loadedTextBoxField);
            }
        }

        // Save the modified PDF with signature fields
        using (FileStream outputFile = new FileStream(Path.GetFullPath(@"../../../Output/Result.pdf"), FileMode.Create))
        {
            loadedDocument.Save(outputFile);
        }
    }
}
#endregion

/// <summary>
/// Represents the Recipient details for the mail merge.
/// </summary>
class Recipient
{
    public string FirstName { get; set; }
    public string CompanyName { get; set; }
    public string Address { get; set; }
    public string PhoneNumber { get; set; }

    public Recipient(string firstName, string companyName, string address, string phone)
    {
        FirstName = firstName;
        CompanyName = companyName;
        Address = address;
        PhoneNumber = phone;
    }
}

