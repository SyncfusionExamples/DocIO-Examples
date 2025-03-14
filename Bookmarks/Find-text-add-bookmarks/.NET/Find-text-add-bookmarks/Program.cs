using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace FindTextAddBookmarks
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                using (WordDocument document = new WordDocument(fileStream, FormatType.Automatic))
                {
                    // Define texts and corresponding bookmark names
                    var textBookmarkPairs = new Dictionary<string, string>
                {
                    { "they are considered one of the world's most loved animals.", "bkmk1" },
                    { "The table below lists the main characteristics the giant panda shares with bears and red pandas.", "bkmk2" },
                    { "Did you know that the giant panda may actually be a raccoon", "bkmk3" }
                };

                    // Add bookmarks to specified texts
                    foreach (var pair in textBookmarkPairs)
                    {
                        AddBookmarkToText(document, pair.Key, pair.Value);
                    }

                    // Retrieve and display bookmark contents
                    List<string> bookmarksContent = GetBookmarkContents(document);
                    foreach (var content in bookmarksContent)
                    {
                        Console.WriteLine("Bookmark content: ");
                        Console.WriteLine(content);
                    }

                    // Save the modified document
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
        /// <summary>
        /// Adds a bookmark to a specific text in the document.
        /// </summary>
        private static void AddBookmarkToText(WordDocument document, string searchText, string bookmarkName)
        {
            TextSelection textSelection = document.Find(searchText, false, true);
            if (textSelection != null)
            {
                WTextRange textRange = textSelection.GetAsOneRange();
                int indexOfText = textRange.OwnerParagraph.Items.IndexOf(textRange);
                textRange.OwnerParagraph.Items.Insert(indexOfText, new BookmarkStart(document, bookmarkName));
                textRange.OwnerParagraph.Items.Insert(indexOfText + 2, new BookmarkEnd(document, bookmarkName));
            }
        }

        /// <summary>
        /// Retrieves all bookmark contents from the document.
        /// </summary>
        private static List<string> GetBookmarkContents(WordDocument document)
        {
            List<string> bookmarkContents = new List<string>();
            foreach (Entity entity in document.FindAllItemsByProperty(EntityType.BookmarkStart, null, null))
            {
                if (entity is BookmarkStart bookmarkStart)
                {
                    var bookmarkNavigator = new BookmarksNavigator(document);
                    bookmarkNavigator.MoveToBookmark(bookmarkStart.Name);
                    WordDocumentPart part = bookmarkNavigator.GetContent();
                    WordDocument tempDoc = part.GetAsWordDocument();
                    bookmarkContents.Add(tempDoc.GetText());

                    tempDoc.Close();
                    tempDoc.Dispose();
                    part.Close();
                }
            }
            return bookmarkContents;
        }
    }
}
