using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

using (FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
{
    // Opens the template Word document.
    using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
    {
        // Find WTextFormField by its name.
        WTextFormField textFormField = document.FindItemByProperty(EntityType.TextFormField, "Name", "Text1") as WTextFormField;
        if (textFormField != null)
        {
            // Get the bookmark name from the textFormField next sibiling.
            string bookmarkName = textFormField.Name;
            // Creates the bookmark navigator instance to access the bookmark.
            BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(document);
            // Moves the virtual cursor to the location before the end of the bookmark.
            bookmarkNavigator.MoveToBookmark(bookmarkName);
            // Gets the bookmark content as WordDocumentPart.
            WordDocumentPart wordDocumentPart = bookmarkNavigator.GetContent();
            // Saves the WordDocumentPart as separate Word document.
            using (WordDocument newDocument = wordDocumentPart.GetAsWordDocument())
            {
                // Close the WordDocumentPart instance.
                wordDocumentPart.Close();
                // Get the text with new line characters.
                string textFieldText = newDocument.GetText();
                // Replace FORMTEXT word from the extracted text.
                textFieldText = textFieldText.Replace(" FORMTEXT ", "");
                // Write the output in console window.
                Console.WriteLine(textFieldText);
                // Press any key in console window to continue.
                Console.ReadKey();
            }
        }
    }
}