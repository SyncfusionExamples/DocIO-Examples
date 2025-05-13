using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

// Load the Word document from the specified path.
using (WordDocument document = new WordDocument(Path.GetFullPath(@"Data/Template.docx")))
{
    // Add bookmarks to various header and footer types in the last section of the document.
    AddBookmarkToHeaderFooter(document, document.LastSection.HeadersFooters.FirstPageHeader, "FirstPageHeader");
    AddBookmarkToHeaderFooter(document, document.LastSection.HeadersFooters.FirstPageFooter, "FirstPageFooter");
    AddBookmarkToHeaderFooter(document, document.LastSection.HeadersFooters.OddHeader, "OddHeader");
    AddBookmarkToHeaderFooter(document, document.LastSection.HeadersFooters.OddFooter, "OddFooter");
    AddBookmarkToHeaderFooter(document, document.LastSection.HeadersFooters.EvenHeader, "EvenHeader");
    AddBookmarkToHeaderFooter(document, document.LastSection.HeadersFooters.EvenFooter, "EvenFooter");

    // Save the modified document to the output path in DOCX format.
    document.Save(Path.GetFullPath(@"Output/Result.docx"), FormatType.Docx);
}

/// <summary>
/// Adds uniquely named bookmarks to paragraphs and table cells that contain content within the given header or footer section.
/// </summary>
void AddBookmarkToHeaderFooter(WordDocument document, HeaderFooter headerFooter, string bookmarkName)
{
    int bookmarkIndex = 1; // Counter to ensure unique bookmark names

    if (headerFooter.ChildEntities.Count > 0)
    {
        foreach (Entity childEntity in headerFooter.ChildEntities)
        {
            if (childEntity is WParagraph paragraph && paragraph.ChildEntities.Count > 0)
            {
                InsertBookmark(document, paragraph, bookmarkName + bookmarkIndex);
                bookmarkIndex++;
            }
            else if (childEntity is WTable table)
            {
                foreach (WTableRow row in table.Rows)
                {
                    foreach (WTableCell cell in row.Cells)
                    {
                        foreach (Entity cellEntity in cell.ChildEntities)
                        {
                            if (cellEntity is WParagraph cellParagraph && cellParagraph.ChildEntities.Count > 0)
                            {
                                InsertBookmark(document, cellParagraph, bookmarkName + bookmarkIndex);
                                bookmarkIndex++;
                            }
                        }
                    }
                }
            }
        }
    }
}

/// <summary>
/// Inserts a bookmark into the given paragraph with the specified name.
/// </summary>
void InsertBookmark(WordDocument document, WParagraph paragraph, string name)
{
    BookmarkStart bookmarkStart = new BookmarkStart(document, name);
    BookmarkEnd bookmarkEnd = new BookmarkEnd(document, name);
    paragraph.ChildEntities.Insert(0, bookmarkStart);
    paragraph.ChildEntities.Add(bookmarkEnd);
}
