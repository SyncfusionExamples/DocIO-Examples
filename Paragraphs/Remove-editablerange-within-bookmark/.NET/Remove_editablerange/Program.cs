using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Remove_editablerange
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Creates a new Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
                {
                    // Loop through all bookmarks in the document
                    for (int i = 0; i < document.Bookmarks.Count; i++)
                    {
                        Bookmark bookmark = document.Bookmarks[i];
                        // Check and remove editable ranges within the bookmark
                        RemoveEditableRange(document, bookmark.Name);
                    }
                    //Create a file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to the file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }

        /// <summary>
        /// Removes any editable ranges found within the bookmark
        /// </summary>
        /// <param name="document"></param>
        /// <param name="bookmarkName"></param>
        private static void RemoveEditableRange(WordDocument document, string bookmarkName)
        {
            // Create a Bookmark Navigator
            BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(document);
            // Move to the bookmark
            bookmarkNavigator.MoveToBookmark(bookmarkName);
            // Get the content inside the bookmark
            WordDocumentPart bookmarkContent = bookmarkNavigator.GetContent();
            // Loop through all sections in the bookmark content
            for (int s = 0; s < bookmarkContent.Sections.Count; s++)
            {
                WSection section = bookmarkContent.Sections[s];
                // Iterate through all entities in the section body (paragraphs, tables, etc.).
                for (int i = 0; i < section.Body.ChildEntities.Count; i++)
                {
                    IEntity entity = section.Body.ChildEntities[i];

                    if (entity is WParagraph)
                    {
                        WParagraph paragraph = entity as WParagraph;
                        // Loop through all child entities in the paragraph
                        for (int j = 0; j < paragraph.ChildEntities.Count; j++)
                        {
                            Entity item = paragraph.ChildEntities[j];
                            // Check if the item is the start of an editable range
                            if (item is EditableRangeStart)
                            {
                                // Find the editable range by ID and remove it
                                EditableRange editableRange = document.EditableRanges.FindById((item as EditableRangeStart).Id);
                                if (editableRange != null)
                                    document.EditableRanges.Remove(editableRange);
                            }
                        }
                    }
                }
            }
        }
    }
}
