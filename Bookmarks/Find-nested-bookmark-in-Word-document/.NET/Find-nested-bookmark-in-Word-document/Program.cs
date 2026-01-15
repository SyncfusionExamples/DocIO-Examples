using Syncfusion.DocIO.DLS;

namespace Find_nested_bookmark_in_Word_document
{
    class Program
    {
        public static void Main(string[] args)
        {
            // Load the existing Word document.
            WordDocument document = new WordDocument(Path.GetFullPath(@"Data/Template.docx"));
            //Iterate through the bookmark collection
            foreach (Bookmark bookmark in document.Bookmarks)
            {
                // Create a collection to store nested bookmarks
                List<string> nestedBookmarks = new List<string>();
                // Create a BookmarksNavigator for the current bookmark
                BookmarksNavigator bookmarksNavigator = new BookmarksNavigator(document);
                {
                    // Move navigator to the current bookmark
                    bookmarksNavigator.MoveToBookmark(bookmark.Name);
                    // Retrieve the bookmark content as a WordDocument
                    using (WordDocument content = bookmarksNavigator.GetContent().GetAsWordDocument())
                    {
                        // Remove headers and footers from the extracted document
                        RemoveHeaderFooter(content);
                        // Iterate through bookmarks in the extracted document
                        foreach (Bookmark bm in content.Bookmarks)
                        {
                            // Exclude the parent itself if names are identical
                            if (!bm.Name.Equals(bookmark.Name))
                                nestedBookmarks.Add(bm.Name);
                        }
                    }
                    // Print the parent bookmark and its nested bookmarks
                    Console.WriteLine("Parent Bookmark: " + bookmark.Name);
                    foreach (string name in nestedBookmarks)
                        Console.WriteLine("Nested Bookmark: " + name);
                    Console.WriteLine("************************");
                }
            }
            Console.ReadLine();
        }
        /// <summary>
        /// Removes all headers and footers from every section in the specified Word document.
        /// </summary>
        /// <param name="doc">The WordDocument instance to process and clear headers and footers.</param>
        /// <remarks>
        /// This method iterates through each section of the document and clears 
        /// the child entities of all header and footer types (first page, odd, even). 
        /// After clearing, it adds an empty paragraph to each header and footer 
        /// to preserve the document’s structure.
        /// </remarks>
        private static void RemoveHeaderFooter(WordDocument doc)
        {
            //Iterate and Remove the Header / footer
            foreach (WSection section in doc.Sections)
            {
                // Remove the first page header and footer
                section.HeadersFooters.FirstPageHeader.ChildEntities.Clear();
                section.HeadersFooters.FirstPageHeader.AddParagraph();
                section.HeadersFooters.FirstPageFooter.ChildEntities.Clear();
                section.HeadersFooters.FirstPageFooter.AddParagraph();

                // Remove the odd page header and footer
                section.HeadersFooters.OddHeader.ChildEntities.Clear();
                section.HeadersFooters.OddHeader.AddParagraph();
                section.HeadersFooters.OddFooter.ChildEntities.Clear();
                section.HeadersFooters.OddFooter.AddParagraph();

                // Remove the even page header and footer
                section.HeadersFooters.EvenHeader.ChildEntities.Clear();
                section.HeadersFooters.EvenHeader.AddParagraph();
                section.HeadersFooters.EvenFooter.ChildEntities.Clear();
                section.HeadersFooters.EvenFooter.AddParagraph();
            }
        }
    }
}