using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;


// Declare a variable to hold the custom character style used for formatting bookmark content.
WCharacterStyle style;

// Load the Word document.
using (WordDocument document = new WordDocument(Path.GetFullPath(@"Data/Template.docx")))
{
    // Navigate to the bookmark named "Tiny_Cubes".
    BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(document);
    bookmarkNavigator.MoveToBookmark("Tiny_Cubes");

    // Extract the content inside the bookmark as a separate Word document.
    WordDocument bookmarkContent = bookmarkNavigator.GetBookmarkContent().GetAsWordDocument();

    // Retrieve the character style named "TinyCube" from the style collection.
    IStyleCollection styleCollection = document.Styles;
    style = styleCollection.FindByName("TinyCube") as WCharacterStyle;

    // Apply the retrieved style to all elements in the extracted bookmark content.
    IterateDocumentElements(bookmarkContent);

    // Create a WordDocumentPart from the modified bookmark content.
    WordDocumentPart wordDocumentPart = new WordDocumentPart(bookmarkContent);

    // Replace the original bookmark content with the styled content.
    bookmarkNavigator.ReplaceContent(wordDocumentPart);

    // Save the updated document to a new file.
    document.Save(Path.GetFullPath(@"Output/Result.docx"), FormatType.Docx);
}

/// <summary>
/// Iterates all sections, headers, and footers in the given document.
/// </summary>
void IterateDocumentElements(WordDocument document)
{
    foreach (WSection section in document.Sections)
    {
        // Process the main body of the section.
        IterateTextBody(section.Body);

        // Process the header and footer (only OddHeader and OddFooter here).
        WHeadersFooters headersFooters = section.HeadersFooters;
        IterateTextBody(headersFooters.OddHeader);
        IterateTextBody(headersFooters.OddFooter);
    }
}

/// <summary>
/// Iterates all entities (paragraphs, tables, block content controls) within a WTextBody.
/// </summary>
void IterateTextBody(WTextBody textBody)
{
    for (int i = 0; i < textBody.ChildEntities.Count; i++)
    {
        IEntity bodyItemEntity = textBody.ChildEntities[i];

        switch (bodyItemEntity.EntityType)
        {
            case EntityType.Paragraph:
                // Process paragraph items (text, fields, etc.)
                IterateParagraph((bodyItemEntity as WParagraph).Items);
                break;

            case EntityType.Table:
                // Recursively process each cell in the table.
                IterateTable(bodyItemEntity as WTable);
                break;

            case EntityType.BlockContentControl:
                // Recursively process the text body within a block content control.
                IterateTextBody((bodyItemEntity as BlockContentControl).TextBody);
                break;
        }
    }
}

/// <summary>
/// Iterates all rows and cells in a table, processing each cell's text body.
/// </summary>
void IterateTable(WTable table)
{
    foreach (WTableRow row in table.Rows)
    {
        foreach (WTableCell cell in row.Cells)
        {
            // Each cell is a TextBody; reuse IterateTextBody to process its content.
            IterateTextBody(cell);
        }
    }
}

/// <summary>
/// Iterates all paragraph items and applies the specified style formatting.
/// </summary>
void IterateParagraph(ParagraphItemCollection paraItems)
{
    for (int i = 0; i < paraItems.Count; i++)
    {
        Entity entity = paraItems[i];

        switch (entity.EntityType)
        {
            case EntityType.TextRange:
                // Apply the character style to the text range.
                (entity as WTextRange).ApplyCharacterFormat(style.CharacterFormat);
                break;

            case EntityType.Field:
                // Apply the character style to the field.
                (entity as WField).ApplyCharacterFormat(style.CharacterFormat);
                break;

            case EntityType.TextBox:
                // Recursively process the contents of the textbox.
                IterateTextBody((entity as WTextBox).TextBoxBody);
                break;

            case EntityType.Shape:
                // Recursively process the contents of the shape.
                IterateTextBody((entity as Shape).TextBody);
                break;

            case EntityType.InlineContentControl:
                // Recursively process the paragraph items within the inline content control.
                IterateParagraph((entity as InlineContentControl).ParagraphItems);
                break;
        }
    }
}