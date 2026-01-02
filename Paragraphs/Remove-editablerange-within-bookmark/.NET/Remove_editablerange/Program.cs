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
            // Get the bookmark content as word document
            WordDocument tempDoc = bookmarkNavigator.GetContent().GetAsWordDocument();
            // Find all entities of type EditableRangeStart within the bookmark
            List<Entity> entity = tempDoc.FindAllItemsByProperty(EntityType.EditableRangeStart, null, null);
            // If any EditableRangeStart entities are found, iterate through them.
            if (entity != null)
            {
                foreach (Entity item in entity)
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
