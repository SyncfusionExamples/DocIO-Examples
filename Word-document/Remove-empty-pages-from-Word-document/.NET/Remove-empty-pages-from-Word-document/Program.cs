using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;


//Opens an existing Word document
using (FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Sample.docx"), FileMode.Open, FileAccess.Read))
{
    using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
    {
        WTextBody textBody = null;
        //Iterates sections in Word document.
        for (int i = document.Sections.Count - 1; i >= 0; i--)
        {
            //Accesses the Body of section where all the contents in document are apart
            textBody = document.Sections[i].Body;
            //Removes the last empty page in the Word document
            RemoveEmptyItems(textBody);
            //Removes the empty sections in the document
            if (textBody.ChildEntities.Count == 0)
            {
                int SectionIndex = document.ChildEntities.IndexOf(document.Sections[i]);
                document.ChildEntities.RemoveAt(SectionIndex);
            }
        }
        using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.Write))
        {
            //Saves and closes the Word document
            document.Save(outputStream, FormatType.Docx);
        }
    }
}
/// <summary>
/// Remove the empty paragraph in the Word document.
/// </summary>
void RemoveEmptyItems(WTextBody textBody)
{
    //A flag to determine any renderable item found in the Word document.
    bool IsRenderableItem = false;
    //A flag to determine a page break is found as the previous item
    bool HasPrevPageBreak = false;
    //Iterates into textbody items.
    for (int itemIndex = textBody.ChildEntities.Count - 1; itemIndex >= 0; itemIndex--)
    {
        //Checks item is empty paragraph and removes it.
        if (textBody.ChildEntities[itemIndex] is WParagraph)
        {
            WParagraph paragraph = textBody.ChildEntities[itemIndex] as WParagraph;
            //Iterates into paragraph
            for (int pIndex = paragraph.Items.Count - 1; pIndex >= 0; pIndex--)
            {
                ParagraphItem paragraphItem = paragraph.Items[pIndex];

                //Removes page breaks
                if ((paragraphItem is Break && (paragraphItem as Break).BreakType == BreakType.PageBreak))
                {
                    if (HasPrevPageBreak)
                        paragraph.Items.RemoveAt(pIndex);
                    else
                        HasPrevPageBreak = true;
                }
                //Check paragraph contains any renderable items.
                else
                {
                    HasPrevPageBreak = false;
                    if (!(paragraphItem is BookmarkStart || paragraphItem is BookmarkEnd))
                    {
                        //Found renderable item and break the iteration.
                        IsRenderableItem = true;
                        break;
                    }
                }
            }
            //Remove empty paragraph and the paragraph with bookmarks only
            if (paragraph.Items.Count == 0 || !IsRenderableItem)
                textBody.ChildEntities.RemoveAt(itemIndex);
        }
    }
}
