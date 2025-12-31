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
                }
            }
            // Retrieve the contents of all bookmarks in the document
            Dictionary<string, string> bookmarkContents = GetBookmarkContents(document);
            // Print each bookmark name and its corresponding content
            foreach (string bkmkName in bookmarkContents.Keys)
            {
                Console.WriteLine("Corresponding Bookmark : " + bkmkName);
                Console.WriteLine("Content : " + bookmarkContents[bkmkName]);
            }
            Console.ReadLine();
        }
        /// <summary>
        /// Retrieves all bookmark contents from the document.
        /// </summary>
        /// <param name="document">The Word document containing bookmarks.</param>
        /// <returns>A dictionary with bookmark names as keys and their text content as values.</returns>
        private static Dictionary<string, string> GetBookmarkContents(WordDocument document)
        {
            // Create a dictionary to store bookmark names and their contents
            Dictionary<string, string> bookmarkContents = new Dictionary<string, string>();
            int paragraphCount = 0;
            // Iterate through each bookmark
            foreach (Bookmark currentBookmark in document.Bookmarks)
            {
                // Create navigator to move to the current bookmark
                BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(document);
                bookmarkNavigator.MoveToBookmark(currentBookmark.Name);
                // Extract the content inside the bookmark as a temporary Word document
                WordDocument tempDoc = bookmarkNavigator.GetContent().GetAsWordDocument();
                // Get the text content and add it to the dictionary
                bookmarkContents.Add(currentBookmark.Name, tempDoc.GetText());
                // Save the Word document.
                tempDoc.Save(Path.GetFullPath("../../../Output/Output_" + paragraphCount + ".docx"));
                // Close the temporary document.
                tempDoc.Close();
                paragraphCount++;
            }
            // Return the dictionary containing all bookmark contents
            return bookmarkContents;
        }
    }
}
