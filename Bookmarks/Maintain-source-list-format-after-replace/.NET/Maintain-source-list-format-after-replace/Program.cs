using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Replace_content_with_bookmark
{
    internal class Program
    {
        static void Main(string[] args)
        {
            using (FileStream destinationStream = new FileStream(Path.GetFullPath("Data/DestinationDocument.docx"), FileMode.Open, FileAccess.Read))
            {
                //Open the destination Word document.
                using (WordDocument destinationDocument = new WordDocument(destinationStream, FormatType.Docx))
                {
                    using (FileStream sourceStream = new FileStream(Path.GetFullPath(@"Data/SourceDocument.docx"), FileMode.Open, FileAccess.Read))
                    {
                        //Open the source Word document.
                        using (WordDocument sourceDocument = new WordDocument(sourceStream, FormatType.Docx))
                        {
                            //Replace text "Text one" in the destination document with content from the bookmark "bkmk1" in the source document.
                            DocxReplaceTextWithDocPart(destinationDocument, sourceDocument, "Text one", "bkmk1");
                            //Replace text "Text two" in the destination document with content from the bookmark "bkmk2" in the source document.
                            DocxReplaceTextWithDocPart(destinationDocument, sourceDocument, "Text two", "bkmk2");
                            //Save the modified destination document to the output stream.
                            using (FileStream output = new FileStream(Path.GetFullPath("Output/Output.docx"), FileMode.Create, FileAccess.Write))
                            {
                                destinationDocument.Save(output, FormatType.Docx);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Replaces specific text in a Word document with bookmarked content from another document, maintaining formatting.
        /// </summary>
        private static void DocxReplaceTextWithDocPart(WordDocument destinationDocument, WordDocument sourceDocument, string tokenToFind, string textBookmark)
        {
            string bookmarkRef = textBookmark + "_bm";
            // Find the text in the destination document where the bookmark start needs to be inserted.
            TextSelection start = destinationDocument.Find(tokenToFind, true, true);
            if (start != null)
            {
                // Get the selected text range and its parent paragraph.
                WTextRange startText = start.GetAsOneRange();
                WParagraph startParagraph = startText.OwnerParagraph;
                // Get the index of the selected text range in the paragraph.
                int index = startParagraph.Items.IndexOf(startText);
                // Remove the selected text at the identified index.
                startParagraph.Items.Remove(startText);
                // Create a BookmarkStart with a unique reference and insert it at the same index.
                BookmarkStart bookmarkStart = new BookmarkStart(destinationDocument, bookmarkRef);
                startParagraph.Items.Insert(index, bookmarkStart);
                // Append a BookmarkEnd with the same reference to mark the bookmark’s end.
                startParagraph.AppendBookmarkEnd(bookmarkRef);

                // Check if the specified bookmark exists in the source document.
                if (sourceDocument.Bookmarks.FindByName(textBookmark) != null)
                {
                    // Move the navigator to the bookmark in the source document.
                    BookmarksNavigator bookmarksNavigator = new BookmarksNavigator(sourceDocument);
                    bookmarksNavigator.MoveToBookmark(textBookmark);
                    // Extract the content within the bookmark.
                    WordDocumentPart wordDocumentPart = bookmarksNavigator.GetContent();

                    // Move the navigator to the newly created bookmark in the destination document.
                    bookmarksNavigator = new BookmarksNavigator(destinationDocument);
                    bookmarksNavigator.MoveToBookmark(bookmarkRef);
                    // Get the paragraph containing the bookmark start in the destination document.
                    WParagraph destinationPara = bookmarksNavigator.CurrentBookmark.BookmarkStart.OwnerParagraph;
                    // Store the list style, first-line indent, and left indent of the paragraph.
                    string listStyleName = destinationPara.ListFormat.CustomStyleName;
                    float firstLineIndent = destinationPara.ParagraphFormat.FirstLineIndent;
                    float leftIndent = destinationPara.ParagraphFormat.LeftIndent;
                    // Replace the bookmark content with the extracted content from the source document.
                    bookmarksNavigator.ReplaceContent(wordDocumentPart);
                    // Reapply the original list style and indent settings to the paragraph.
                    destinationPara.ListFormat.ApplyStyle(listStyleName);
                    destinationPara.ParagraphFormat.FirstLineIndent = firstLineIndent;
                    destinationPara.ParagraphFormat.LeftIndent = leftIndent;
                }
                else
                {
                    // If the bookmark is not found, replace the content with an empty string.
                    BookmarksNavigator bookmarksNavigator = new BookmarksNavigator(destinationDocument);
                    bookmarksNavigator.MoveToBookmark(bookmarkRef);
                    bookmarksNavigator.ReplaceBookmarkContent(string.Empty, true);
                }
            }
        }
    }
}
