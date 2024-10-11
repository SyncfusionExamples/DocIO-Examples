using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

// Open the existing main document from a file stream.
using (FileStream destinationStream = new FileStream(Path.GetFullPath(@"Data/DestinationDocument.docx"), FileMode.Open))
{
    // Load the main Word document.
    WordDocument destinationDocument = new WordDocument(destinationStream, FormatType.Docx);

    // Open the existing source document from a file stream.
    using (FileStream sourceStream = new FileStream(Path.GetFullPath(@"Data/SourceDocument.docx"), FileMode.Open))
    {
        // Load the source Word document.
        WordDocument sourceDocument = new WordDocument(sourceStream, FormatType.Docx);

        // Set the first section break in the source document to "NoBreak" for seamless merging.
        sourceDocument.Sections[0].BreakCode = SectionBreakCode.NoBreak;

        // Get the index of the last section in the destination document.
        int secIndex = destinationDocument.ChildEntities.IndexOf(destinationDocument.LastSection);

        // Get the index of the last paragraph in the last section of the destination document.
        int paraIndex = destinationDocument.LastSection.Body.ChildEntities.IndexOf(destinationDocument.LastParagraph);

        // Get the style and formatting of the last paragraph for reference.
        WParagraph lastPara = destinationDocument.LastParagraph;

        // Import content from the source document into the destination document, using destination styles.
        destinationDocument.ImportContent(sourceDocument, ImportOptions.UseDestinationStyles);

        // Modify the paragraph style for the newly added contents by applying left indentation.
        AddLeftIndentation(destinationDocument, secIndex, paraIndex + 1, lastPara.ParagraphFormat.LeftIndent);

        // Save the updated destination document to a new file.
        using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.Write))
        {
            destinationDocument.Save(outputStream, FormatType.Docx);
        };
    };
};

/// <summary>
/// Applies left indentation to paragraphs and tables in a specified section and paragraph range of a Word document.
/// </summary>
void AddLeftIndentation(WordDocument document, int secIndex, int paraIndex, float leftIndent)
{
    // Iterate through the sections added from the source document, starting from secIndex.
    for (int i = secIndex; i < document.ChildEntities.IndexOf(document.LastSection) + 1; i++)
    {
        // Iterate through the child entities (paragraphs/tables) added from the source document.
        for (int j = paraIndex; j < document.Sections[i].Body.ChildEntities.Count; j++)
        {
            // If the child entity is a paragraph, apply the left indent from the last paragraph.
            if (document.Sections[i].Body.ChildEntities[j] is WParagraph)
            {
                WParagraph para = document.Sections[i].Body.ChildEntities[j] as WParagraph;
                // Set the left indentation for the paragraph
                para.ParagraphFormat.LeftIndent = leftIndent;
            }
            // If the child entity is a table, apply the same left indent to the table.
            else if (document.Sections[i].Body.ChildEntities[j] is WTable)
            {
                WTable table = document.Sections[i].Body.ChildEntities[j] as WTable;
                // Set the left indentation for the table.
                table.TableFormat.LeftIndent = leftIndent;
            }
        }
        // Reset the paragraph index for the next section to start from the beginning.
        paraIndex = 0;
    }
}