using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

using (FileStream destinationStream = new FileStream(Path.GetFullPath("Data/DestinationDocument.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    // Open the destination Word document.
    using (WordDocument destinationDocument = new WordDocument(destinationStream, FormatType.Docx))
    {
        using (FileStream sourceStream = new FileStream(Path.GetFullPath("Data/SourceDocument.docx"), FileMode.Open, FileAccess.ReadWrite))
        {
            // Open the source Word document.
            using (WordDocument sourceDocument = new WordDocument(sourceStream, FormatType.Docx))
            {
                // Set the break code as no break to prevent page breaks.
                sourceDocument.Sections[0].BreakCode = SectionBreakCode.NoBreak;
                // Initialize a variable to hold the list style from the destination document.
                ListStyle listStyle = null;
                // Iterate through the paragraphs in the last section of the destination document.
                foreach (WParagraph paragraph in destinationDocument.LastSection.Paragraphs)
                {
                    if (paragraph.ListFormat.CurrentListStyle != null)
                    {
                        // Get the current list style.
                        listStyle = paragraph.ListFormat.CurrentListStyle;
                    }
                    else
                    {
                        // Check the paragraph style for a list style and store it.
                        WParagraphStyle style = destinationDocument.Styles.FindByName(paragraph.StyleName) as WParagraphStyle;
                        if (style != null)
                            listStyle = style.ListFormat.CurrentListStyle;
                    }
                    // Break the loop if a list style is found.
                    if (listStyle != null)
                        break;
                }
                // Import the content of the source document at the end of the destination document.
                destinationDocument.ImportContent(sourceDocument, ImportOptions.ListContinueNumbering);
                if (listStyle != null)
                {
                    // Apply liststyle to the paragraphs in the destination document to maintain continuous numbering.
                    foreach (WParagraph paragraph in destinationDocument.LastSection.Paragraphs)
                    {
                        if (paragraph.ListFormat.CurrentListStyle != null)
                            paragraph.ListFormat.ApplyStyle(listStyle.Name);
                    }
                }
                // Save the merged document.
                using (FileStream outputStream = new FileStream(Path.GetFullPath("Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    destinationDocument.Save(outputStream, FormatType.Docx);
                }
            }
        }
    }
}
