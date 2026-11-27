using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Collections.Generic;
using System.IO;

namespace Remove_contents_after_specific_text
{
    class Program
    {
        static void Main(string[] args)
        {
            // Open the input Word document
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                // Load the existing Word document into the DocIO instance.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Automatic))
                {
                    // Find all TextRange entities with the text "ExampleText".
                    List<Entity> textRange = document.FindAllItemsByProperty(EntityType.TextRange, "Text", "ExampleText");
                    // Remove the content that occurs after the first matching text range.
                    RemoveContenAfterTextRange(document, textRange[0] as WTextRange);
                    // Save the modified document to the specified output path.
                    document.Save(@"../../../Output/output.docx");
                }
            }
        }
        /// <summary>
        /// Removes content after a specified WTextRange in the Word document.
        /// </summary>
        /// <param name="document">The WordDocument instance.</param>
        /// <param name="txtRange">The WTextRange after which content should be removed.</param>
        private static void RemoveContenAfterTextRange(WordDocument document, WTextRange txtRange)
        {
            // Get the parent paragraph of the input text range.
            WParagraph ownerPara = txtRange.OwnerParagraph;
            // Create a start bookmark immediately after the target text range.
            BookmarkStart start = new BookmarkStart(document, "customBookmark");
            int textRangeIndex = ownerPara.ChildEntities.IndexOf(txtRange);
            ownerPara.ChildEntities.Insert(++textRangeIndex, start);
            // Create an end bookmark at the end of the document.
            BookmarkEnd end = new BookmarkEnd(document, "customBookmark");
            document.LastParagraph.ChildEntities.Add(end);
            // Use BookmarksNavigator to select and delete all content between the start and end bookmarks.
            BookmarksNavigator bkNavigator = new BookmarksNavigator(document);
            bkNavigator.MoveToBookmark("customBookmark");
            bkNavigator.DeleteBookmarkContent(true, true); // Removes bookmark content and the bookmark itself.
        }
    }
}