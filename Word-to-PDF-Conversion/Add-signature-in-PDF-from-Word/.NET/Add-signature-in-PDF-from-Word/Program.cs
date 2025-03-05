using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Interactive;
using Syncfusion.Pdf.Parsing;
using Syncfusion.Drawing;

using (FileStream docStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read))
{
    using (WordDocument document = new WordDocument(docStream, FormatType.Docx))
    {
        using (DocIORenderer renderer = new DocIORenderer())
        {
            //Sets true to preserve the Word document form field as editable PDF form field in PDF document.
            renderer.Settings.PreserveFormFields = true;

            using (PdfDocument pdfDocument = renderer.ConvertToPDF(document))
            {
                using (MemoryStream stream = new MemoryStream())
                {
                    pdfDocument.Save(stream);
                    stream.Position = 0;
                    //Add signature field in PDF.
                    AddSignature(stream);
                }
            }
        }
    }
}

/// <summary>
/// Adds signature field in PDF
/// </summary>
/// <param name="stream"></param>
static void AddSignature(MemoryStream stream)
{
    //Load the PDF document.
    PdfLoadedDocument loadedDocument = new PdfLoadedDocument(stream);
    //Get the loaded form.
    PdfLoadedForm loadedForm = loadedDocument.Form;
    for (int i = 0; i < loadedForm.Fields.Count; i++)
    {
        if (loadedForm.Fields[i] is PdfLoadedTextBoxField)
        {
            //Get the loaded text box field and fill it.
            PdfLoadedTextBoxField loadedTextBoxField = loadedForm.Fields[i] as PdfLoadedTextBoxField;
            //Get bounds from an existing textbox field.
            RectangleF bounds = loadedTextBoxField.Bounds;
            //Get page.
            PdfPageBase loadedPage = loadedTextBoxField.Page;
            //Create PDF Signature field.
            PdfSignatureField signatureField = new PdfSignatureField(loadedPage, loadedTextBoxField.Text.Trim());
            //Set properties to the signature field.
            signatureField.Bounds = bounds;
            //Add the form field to the document.
            loadedDocument.Form.Fields.Add(signatureField);
        }
    }
    //Save the document.
    using (FileStream outputStream1 = new FileStream(Path.GetFullPath(@"Output/Result.pdf"), FileMode.OpenOrCreate, FileAccess.ReadWrite))
    {
        loadedDocument.Save(outputStream1);
    }
}