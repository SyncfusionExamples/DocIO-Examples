using Syncfusion.DocIO.DLS;

namespace Delete_bracketed_content_up_to_specific_word
{
    class Program
    {
        public static void Main(string[] args)
        {
            // Load the existing word document
            using(WordDocument document = new WordDocument(@"Data\Input.docx"))
            {
                // Find the phrase using Find API (Specific Word, case-insensitive, whole-word match)
                TextSelection selection = document.Find("Importadores", true, true);
                // Proceed only if the phrase is found
                if (selection != null)
                {
                    // Get the paragraph that contains the selected word
                    WParagraph paragraph = selection.GetAsOneRange().OwnerParagraph;
                    // Get a owner section
                    WSection ownerSection = (WSection)paragraph.OwnerTextBody.Owner;
                    // Find the position (index) of the paragraph in the section
                    int phraseParaIndex = ownerSection.Body.ChildEntities.IndexOf(paragraph);
                    // Find the position of the word inside the paragraph
                    int matchWordIndex = paragraph.ChildEntities.IndexOf(selection.GetAsOneRange());
                    // Set how many paragraphs before the word we want to check
                    int maxPreviousParagraphs = 6;
                    // Call a method to remove content inside brackets before the word
                    RemoveBlock(document, paragraph, phraseParaIndex, matchWordIndex, ownerSection, maxPreviousParagraphs);
                    // Save the updated document to a file
                    document.Save(@"../../../Output/Output.docx");
                }
            }
        }
        /// <summary>
        /// Removes content enclosed in brackets from previous paragraphs up to a specific word in the document.
        /// </summary>
        /// <param name="document">The Word document to modify.</param>
        /// <param name="paragraph">The paragraph that contains the specific word or phrase.</param>
        /// <param name="phraseParaIndex">The index of the paragraph containing the specific word.</param>
        /// <param name="matchWordIndex">The index of the word within the paragraph.</param>
        /// <param name="ownerSection">The section that contains the paragraph.</param>
        /// <param name="maxPreviousParagraphs">The number of previous paragraphs to check for bracketed content.</param>
        private static void RemoveBlock(WordDocument document, WParagraph paragraph, int phraseParaIndex, int matchWordIndex, WSection ownerSection, int maxPreviousParagraphs)
        {
            // Initialize state
            int bracketCount = 0;
            bool isBracketFound = false;
            // Create a unique name for the temporary bookmark
            string bookmarkName = "Remove" + Guid.NewGuid().ToString();
            // Store the paragraph where the opening bracket '[' is found
            WParagraph openingBracketParagraph = null;
            // Store the character position of the opening bracket '['
            int openingBracketCharIndex = 0;
            // Loop backwards through previous paragraphs
            for (int i = phraseParaIndex; i >= 0 && i > phraseParaIndex - maxPreviousParagraphs; i--)
            {
                WParagraph currentParagraph = ownerSection.Paragraphs[i];
                // If it's the phrase paragraph, start from the word's position; otherwise, start from the end
                int start = i == phraseParaIndex ? matchWordIndex : currentParagraph.ChildEntities.Count - 1;
                // Loop through entities in reverse
                for (int j = start; j >= 0; j--)
                {
                    // Check if the entity is a text range
                    if (currentParagraph.ChildEntities[j] is WTextRange textRange)
                    {
                        string text = textRange.Text;
                        // Loop through characters in reverse
                        for (int k = text.Length - 1; k >= 0; k--)
                        {
                            char ch = text[k];
                            // If we find a closing bracket
                            if (ch == ']')
                            {
                                isBracketFound = true;
                                bracketCount++;
                            }
                            // If we find an opening bracket
                            else if (ch == '[')
                            {
                                // If there's no matching closing bracket, stop processing
                                if (bracketCount == 0)
                                    return;

                                // Reduce bracket count
                                bracketCount--;
                                // Save the position of the opening bracket
                                openingBracketCharIndex = k;
                                // Store the paragraph that contains this opening bracket,
                                openingBracketParagraph  = currentParagraph;
                            }
                        }
                    }
                }
            }
            // When a bracket is found and bracket count is 0, it means a complete and valid bracket pair is detected.
            if (isBracketFound && bracketCount == 0)
            {
                // Create and insert a bookmark start at the opening bracket position
                BookmarkStart bookmarkStart = new BookmarkStart(document, bookmarkName);
                openingBracketParagraph .ChildEntities.Insert(openingBracketCharIndex, bookmarkStart);
                // Create and insert a bookmark end after the specific word
                BookmarkEnd bookmarkEnd = new BookmarkEnd(document, bookmarkName);
                paragraph.ChildEntities.Insert(matchWordIndex + 1, bookmarkEnd);
                // Use Bookmarknavigator to delete content between the bookmarks
                BookmarksNavigator navigator = new BookmarksNavigator(document);
                navigator.MoveToBookmark(bookmarkName);
                navigator.DeleteBookmarkContent(true);
                // Remove the temporary bookmark from the document
                document.Bookmarks.Remove(navigator.CurrentBookmark);
            }
        }
    }
}
