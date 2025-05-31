using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;



// Define the name of the first existing bookmark from where extraction starts.
string bookmark1 = "Mysteries";
// Define the name of the second existing bookmark where extraction ends.
string bookmark2 = "Facts";
// Define the name for a temporary bookmark that spans content between bookmark1 and bookmark2.
string tempBookmarkName = "tempBookmark";

// Load the Word document.
using (WordDocument document = new WordDocument(Path.GetFullPath(@"Data/Template.docx")))
{
    // Find the first bookmark by name.
    Bookmark firstBookmark = document.Bookmarks.FindByName(bookmark1);
    int index = 0;

    // Get the paragraph that contains the end of the first bookmark.
    WParagraph firstBookmarkOwnerPara = firstBookmark.BookmarkEnd.OwnerParagraph;
    // Find the position of the bookmark end in the paragraph.
    index = firstBookmarkOwnerPara.Items.IndexOf(firstBookmark.BookmarkEnd);

    // Create a temporary bookmark start and insert it right after the first bookmark end.
    BookmarkStart newBookmarkStart = new BookmarkStart(document, tempBookmarkName);
    firstBookmarkOwnerPara.ChildEntities.Insert(index + 1, newBookmarkStart);

    // Find the second bookmark by name.
    Bookmark secondBookmark = document.Bookmarks.FindByName(bookmark2);
    // Get the paragraph that contains the start of the second bookmark.
    WParagraph secondBookmarkOwnerPara = secondBookmark.BookmarkStart.OwnerParagraph;
    // Find the position of the bookmark start in the paragraph.
    index = secondBookmarkOwnerPara.Items.IndexOf(secondBookmark.BookmarkStart);

    // Create a temporary bookmark end and insert it just before the second bookmark start.
    BookmarkEnd newBookmarkEnd = new BookmarkEnd(document, tempBookmarkName);
    secondBookmarkOwnerPara.ChildEntities.Insert(index, newBookmarkEnd);

    // Navigate to the temporary bookmark created between the two bookmarks.
    BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(document);
    bookmarkNavigator.MoveToBookmark(tempBookmarkName);

    // Extract the content between the bookmarks as TextBodyPart (optional if needed later).
    TextBodyPart part = bookmarkNavigator.GetBookmarkContent();

    // Get the bookmark content as a new Word document part.
    WordDocumentPart content = bookmarkNavigator.GetContent();

    // Load the extracted content into a temporary Word document for modification or export.
    using (WordDocument tempDocument = content.GetAsWordDocument())
    {
        // Remove all headers and footers from the extracted content.
        RemoveHeaderAndFooter(tempDocument);

        // Save the extracted content as an HTML file.
        tempDocument.Save(Path.GetFullPath(@"Output/Result.html"), FormatType.Html);
    }
}

// Removes all headers and footers from the given Word document.
void RemoveHeaderAndFooter(WordDocument document)
{
    foreach (WSection section in document.Sections)
    {
        foreach (HeaderFooter entity in section.HeadersFooters)
        {
            if (entity != null)
                entity.ChildEntities.Clear(); 
        }
    }
}

