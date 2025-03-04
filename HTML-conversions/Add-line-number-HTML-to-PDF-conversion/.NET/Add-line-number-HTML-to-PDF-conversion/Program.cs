using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;

// Open the HTML file as a stream.
using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Sample.html"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    // Load the file stream into a Word document in HTML format.
    using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Html))
    {
        // Instantiate DocIORenderer to handle Word to PDF conversion.
        DocIORenderer render = new DocIORenderer();

        // Split the HTML content into sections and add line numbers where necessary.
        SplitSectionAndSetLineNumber(document);

        // Convert the modified Word document into a PDF document.
        PdfDocument pdfDocument = render.ConvertToPDF(document);

        // Create a file stream for saving the output PDF file.
        using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output.pdf"), FileMode.Create, FileAccess.ReadWrite))
        {
            // Save the generated PDF document to the output file stream.
            pdfDocument.Save(outputFileStream);
        }
    }
}

/// <summary>
/// Splits the document content into sections at a placeholder and enables line numbering.
/// </summary>
/// <param name="document">The Word document to modify.</param>
void SplitSectionAndSetLineNumber(WordDocument document)
{
    // Get the first and only section of the document.
    WSection section = document.Sections[0];

    // Find the text range for the placeholder text ("Page2").
    WTextRange placeHolder = document.Find("Page2", true, true).GetAsOneRange();

    // Get the index of the paragraph containing the placeholder.
    int index = section.Body.ChildEntities.IndexOf(placeHolder.OwnerParagraph);

    // Clone the entire section to create a new section starting from the placeholder.
    WSection clonedSection = section.Clone();

    // Remove all content before the placeholder paragraph in the cloned section.
    for (int i = index - 1; i >= 0; i--)
        clonedSection.Body.ChildEntities.RemoveAt(i);

    // Add the modified cloned section to the document.
    document.ChildEntities.Add(clonedSection);

    // Remove all content from the placeholder paragraph to the end in the original section.
    for (int i = index; i < section.Body.ChildEntities.Count;)
        section.Body.ChildEntities.RemoveAt(i);

    // Enable line numbering to restart at each section.
    section.PageSetup.LineNumberingMode = LineNumberingMode.RestartSection;
    clonedSection.PageSetup.LineNumberingMode = LineNumberingMode.RestartSection;
}
