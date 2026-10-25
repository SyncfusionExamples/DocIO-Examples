using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;


//Creates a new Word document
WordDocument document = new WordDocument(@"../../../Data/Template.docx");
//Updates the table of contents
document.UpdateTableOfContents();
TableOfContent tableOfContents = (document.Sections[0].Paragraphs[0] as WParagraph).ChildEntities[0] as TableOfContent;
//Iterate the toc section and apply line spacing.
ApplyLineSpacingToTOC(tableOfContents.OwnerParagraph.OwnerTextBody);
//Saves and closes the Word document instance
document.Save(@"../../../Output/Result.docx", FormatType.Docx);
document.Close();

/// <summary>
/// Apply Line spacing to TOC.
/// </summary>
/// <param name="textBody"></param>
void ApplyLineSpacingToTOC(WTextBody textBody)
{
    for (int i = 0; i < textBody.ChildEntities.Count; i++)
    {
        if (textBody.ChildEntities[i] is WParagraph)
        {
            if (CheckTOCEntries((textBody.ChildEntities[i] as WParagraph).Items))
            {
                (textBody.ChildEntities[i] as WParagraph).ParagraphFormat.LineSpacing = 30;
            }
        }
    }

}
/// <summary>
/// Iterates into paragraph items.
/// </summary>
/// <param name="paragraph">The paragraph.</param>
/// <param name="fieldType">Type of field.</param>
bool CheckTOCEntries(ParagraphItemCollection paraItems)
{
    for (int i = 0; i < paraItems.Count; i++)
    {
        if (paraItems[i] is WField && (paraItems[i] as WField).FieldType == FieldType.FieldHyperlink)
        {
            return true;
        }
    }
    return false;
}