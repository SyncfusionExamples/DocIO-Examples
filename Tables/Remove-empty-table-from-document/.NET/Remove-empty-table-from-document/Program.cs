using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;


// Load the Word document
using (WordDocument document = new WordDocument(Path.GetFullPath(@"Data/Template.docx")))
{
    // Iterate through all sections in the document
    foreach (WSection section in document.Sections)
    {
        // Iterate through tables in reverse order to safely remove them if needed
        for (int i = section.Tables.Count - 1; i >= 0; i--)
        {
            WTable table = (WTable)section.Tables[i];

            #region RemoveNestedTableFirst
            // Iterate through rows of the table
            foreach (WTableRow row in table.Rows)
            {
                // Iterate through cells of the row
                foreach (WTableCell cell in row.Cells)
                {
                    // Iterate through child entities in reverse order for safe removal
                    for (int j = cell.ChildEntities.Count - 1; j >= 0; j--)
                    {
                        Entity entity = cell.ChildEntities[j];

                        // Check if entity is a table and if it's completely empty
                        if (entity.EntityType == EntityType.Table && IsTableCompletelyEmpty(entity as WTable))
                        {
                            cell.ChildEntities.Remove(entity); // Remove empty nested table
                        }
                    }
                }
            }
            #endregion

            // If the entire table is empty, remove it from the section
            if (IsTableCompletelyEmpty(table))
            {
                section.Body.ChildEntities.Remove(table);
            }
        }
    }
    // Save the modified document
    document.Save(Path.GetFullPath(@"Output/Result.docx"), FormatType.Docx);
}  

/// <summary>
/// Checks whether table is empty
/// </summary>
/// <param name="table"></param>
/// <returns>True if table is empty, Otherwise False.</returns>
static bool IsTableCompletelyEmpty(WTable table)
{
    for (int i = 0; i < table.Rows.Count; i++)
    {
        WTableRow row = table.Rows[i];

        for (int j = 0; j < row.Cells.Count; j++)
        {
            WTableCell cell = row.Cells[j];
            // If any cell contains content, the table is not empty
            if (!IsTextBodyEmpty(cell))
                return false;
        }
    }
    return true;
}
/// <summary>
/// Checks whether text body is empty
/// </summary>
/// <param name="textBody"></param>
/// <returns></returns>
static bool IsTextBodyEmpty(WTextBody textBody)
{
    for (int i = textBody.ChildEntities.Count - 1; i >= 0; i--)
    {
        Entity entity = textBody.ChildEntities[i];

        switch (entity.EntityType)
        {
            case EntityType.Paragraph:
                if (!IsParagraphEmpty(entity as WParagraph))
                    return false;
                break;
            case EntityType.BlockContentControl:
                if (!IsTextBodyEmpty((entity as BlockContentControl).TextBody))
                    return false;
                break;
        }
    }
    return true;
}
/// <summary>
/// Checks whether paragraph is empty
/// </summary>
/// <param name="textBody"></param>
/// <returns></returns>
static bool IsParagraphEmpty(WParagraph paragraph)
{
    for (int i = 0; i < paragraph.ChildEntities.Count; i++)
    {
        Entity entity = paragraph.ChildEntities[i];
        switch (entity.EntityType)
        {
            case EntityType.TextRange:
                WTextRange textRange = entity as WTextRange;
                if (!string.IsNullOrEmpty(textRange.Text))
                    return false;
                break;
            case EntityType.BookmarkStart:
            case EntityType.BookmarkEnd:
            case EntityType.EditableRangeStart:
            case EntityType.EditableRangeEnd:
                break;
            default:
                return false;
        }
    }
    return true;
}
