using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Replace_content_with_bookmark
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Open the source Word document for reading using a file stream.
            using (FileStream sourceStream = new FileStream(Path.GetFullPath("Data/SourceDocument.docx"), FileMode.Open, FileAccess.Read))
            {
                // Load the source document from the stream.
                using (WordDocument sourceDocument = new WordDocument(sourceStream, FormatType.Docx))
                {
                    // Open the destination Word document for reading using another file stream.
                    using (FileStream destinationStream = new FileStream(Path.GetFullPath(@"Data/DestinationDocument.docx"), FileMode.Open, FileAccess.Read))
                    {
                        // Load the destination document from the stream.
                        using (WordDocument destinationDocument = new WordDocument(destinationStream, FormatType.Docx))
                        {
                            // Replace text "Text one" in the source document with content from the bookmark "bkmk1" in the destination document.
                            DocxReplaceTextWithDocPart(sourceDocument, destinationDocument, "Text one", "bkmk1");
                            // Replace text "Text two" in the source document with content from the bookmark "bkmk2" in the destination document.
                            DocxReplaceTextWithDocPart(sourceDocument, destinationDocument, "Text two", "bkmk2");
                            // Open a stream to save the modified source document as a new file.
                            using (FileStream output = new FileStream(Path.GetFullPath("Output/Output.docx"), FileMode.Create, FileAccess.Write))
                            {
                                // Save the modified source document to the output stream in DOCX format.
                                sourceDocument.Save(output, FormatType.Docx);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Replaces specific text in a Word document with bookmarked content from another document, maintaining formatting.
        /// </summary>
        private static void DocxReplaceTextWithDocPart(WordDocument document, WordDocument sourceDoc, string tokenToFind, string textBookmark)
        {
            string bookmarkRef = textBookmark + "_bm";

            // Find the start token in the document.
            TextSelection start = document.Find(tokenToFind, true, true);
            if (start != null)
            {
                WTextRange startText = start.GetAsOneRange();
                WParagraph startParagraph = startText.OwnerParagraph;
                // Get the index of the start token in the paragraph.
                int index = startParagraph.Items.IndexOf(startText);
                // Remove the start token at the specified index.
                startParagraph.Items.Remove(startText);
                // Create and insert a BookmarkStart at the index of the start token.
                BookmarkStart bookmarkStart = new BookmarkStart(document, bookmarkRef);
                startParagraph.Items.Insert(index, bookmarkStart);
                startParagraph.AppendBookmarkEnd(bookmarkRef);

                // Check if the bookmark exists in the source document.
                if (sourceDoc.Bookmarks.FindByName(textBookmark) != null)
                {
                    // Access the bookmark in the source document.
                    BookmarksNavigator bookmarksNavigator = new BookmarksNavigator(sourceDoc);
                    bookmarksNavigator.MoveToBookmark(textBookmark);
                    // Get the bookmark content.
                    WordDocumentPart wordDocumentPart = bookmarksNavigator.GetContent();
                    bookmarksNavigator = new BookmarksNavigator(document);
                    bookmarksNavigator.MoveToBookmark(bookmarkRef);

                    // Get the destination paragraph before replacing.
                    WParagraph destinationPara = bookmarksNavigator.CurrentBookmark.BookmarkStart.OwnerParagraph;
                    // Store the list style, first line indent, and left indent of the paragraph.
                    string listStyleName = destinationPara.ListFormat.CustomStyleName;
                    float firstLineIndent = destinationPara.ParagraphFormat.FirstLineIndent;
                    float leftIndent = destinationPara.ParagraphFormat.LeftIndent;

                    // Replace the selected text with the bookmark content from the source document.
                    bookmarksNavigator.ReplaceContent(wordDocumentPart);
                    // Reapply the list style and indent values after replacement.
                    destinationPara.ListFormat.ApplyStyle(listStyleName);
                    destinationPara.ParagraphFormat.FirstLineIndent = firstLineIndent;
                    destinationPara.ParagraphFormat.LeftIndent = leftIndent;
                }
                else
                {
                    // If the bookmark is not found, replace the bookmark content with an empty string.
                    BookmarksNavigator bookmarksNavigator = new BookmarksNavigator(document);
                    bookmarksNavigator.MoveToBookmark(bookmarkRef);
                    bookmarksNavigator.ReplaceBookmarkContent(string.Empty, true);
                }
            }
        }

    }
}
