using Syncfusion.DocIO.DLS;

namespace Find_Bookmark_Owner_TextBody_or_Header_Footer
{
    class Program
    {
        public static void Main(string[] args)
        {
            // Load the existing Word document
            WordDocument document = new WordDocument(Path.GetFullPath("Data/Template.docx"));
            // Iterate through all bookmarks in the document.
            foreach (Bookmark bookmark in document.Bookmarks)
            {
                // Get bookmark start from the current bookmark
                BookmarkStart bkmkStart = bookmark.BookmarkStart;
                if (bkmkStart != null)
                {
                    // Get the paragraphs that contain the bookmark's start.
                    WParagraph startPara = bkmkStart.OwnerParagraph;
                    if (startPara == null)
                        continue;

                    Entity ownerEntity = startPara;
                    if (ownerEntity != null)
                    {
                        // Traverse the owner hierarchy until reaching the section, stopping if a HeaderFooter is found
                        while (!(ownerEntity is WSection))
                        {
                            if (ownerEntity.EntityType == EntityType.HeaderFooter)
                                break;
                            ownerEntity = ownerEntity.Owner;
                        }
                    }
                    // Check if the bookmark is in the text body, header, or footer
                    string ownerLabel = (ownerEntity.EntityType == EntityType.Section)
                        ? "TextBody"
                        : CheckHeaderFooterType(ownerEntity.Owner as WSection, ownerEntity as HeaderFooter);
                    // Print the bookmark name and its owner type
                    Console.WriteLine("Bookmark Name:" + bkmkStart.Name + "\n  Bookmark Owner:" + ownerLabel);                    
                }
            }
            Console.ReadLine();
        }
        /// <summary>
        /// Returns a whether the provided HeaderFooter instance belongs to the header or footer of the given section.
        /// </summary>
        /// <param name="section">The section that contains the HeaderFooter</param>
        /// <param name="headerFooter">The HeaderFooter instance to check.</param>
        /// <returns>Returns "Header" if the instance is a header, "Footer" if it is a footer,
        /// otherwise "Header and Footer".</returns>
        private static string CheckHeaderFooterType(WSection section, HeaderFooter headerFooter)
        {
            string type = "Header and Footer";
            // Check if the given HeaderFooter instance is one of the section's header references.
            if (section.HeadersFooters.OddHeader == headerFooter
                || section.HeadersFooters.FirstPageHeader == headerFooter || section.HeadersFooters.EvenHeader == headerFooter)
            {
                type = "Header";
            }
            // Otherwise, check if it is one of the section's footer references.
            else if (section.HeadersFooters.OddFooter == headerFooter || section.HeadersFooters.EvenFooter == headerFooter
                || section.HeadersFooters.FirstPageFooter == headerFooter)
            {
                type = "Footer";
            }
            return type;
        }
    }
}
