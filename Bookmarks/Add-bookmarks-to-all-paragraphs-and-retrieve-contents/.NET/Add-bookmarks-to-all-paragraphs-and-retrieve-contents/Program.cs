using Syncfusion.DocIO.DLS;

namespace Add_bookmarks_to_all_paragraphs_and_retrieve_contents
{
    class Program
    {
        public static void Main(string[] args)
        {
            // Load the existing Word document
            WordDocument document = new WordDocument(Path.GetFullPath("Data/Input.docx"));
            // Retrieve all paragraph entities in the document
            List<Entity> paragraphsToInsertBookmarks = document.FindAllItemsByProperty(EntityType.Paragraph, null, null);
            foreach (Entity entity in paragraphsToInsertBookmarks)
            {
                // Cast the entity to a paragraph
                WParagraph currentPara = entity as WParagraph;
                // Skip the paragraph if it is empty
                if (currentPara.Text != string.Empty)
                {
                    // Create a unique bookmark name using a GUID
                    string bookmarkName = "Bookmark" + Guid.NewGuid();
                    // Insert a bookmark start at the beginning of the paragraph
                    currentPara.ChildEntities.Insert(0, new BookmarkStart(document, bookmarkName));
                    // Insert a bookmark end at the end of the paragraph
                    currentPara.AppendBookmarkEnd(bookmarkName);
                    // Print the bookmark name and the paragraph text to the console
                    Console.WriteLine("Corresponding Bookmark : " + bookmarkName);
                    Console.WriteLine("Content : " + currentPara.Text + "\n");
                }
            }
            // Save each bookmarked paragraph as a separate document
            SaveBookmarksAsSeparateDocuments(document);
            // Close the Word document
            document.Close();
            Console.ReadLine();
        }
        /// <summary>
        /// Saves the content of each bookmark in the given Word document as a separate Word document file.
        /// </summary>
        /// <param name="document">The Word document containing bookmarks.</param>
        private static void SaveBookmarksAsSeparateDocuments(WordDocument document)
        {
            int paragraphCount = 0;
            // Iterate through each bookmark
            foreach (Bookmark currentBookmark in document.Bookmarks)
            {
                // Create navigator to move to the current bookmark
                BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(document);
                bookmarkNavigator.MoveToBookmark(currentBookmark.Name);
                // Extract the content inside the bookmark as a temporary Word document
                WordDocument tempDoc = bookmarkNavigator.GetContent().GetAsWordDocument();
                // Save the Word document.
                tempDoc.Save(Path.GetFullPath("../../../Output/Output_" + paragraphCount + ".docx"));
                // Close the temporary document.
                tempDoc.Close();
                paragraphCount++;
            }
        }
    }
}
