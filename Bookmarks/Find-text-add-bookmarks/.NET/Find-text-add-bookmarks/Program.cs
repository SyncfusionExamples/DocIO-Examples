using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Find_text_add_bookmarks
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
            {
                //Loads an existing Word document into DocIO instance.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Automatic))
                {
                    //Finds the first occurrence of a particular text in the document
                    TextSelection textSelection = document.Find("they are considered one of the world's most loved animals.", false, true);
                    //Gets the found text as single text range
                    WTextRange textRange = textSelection.GetAsOneRange();
                    //Add bookmark to the selected text range.
                    //Create bookmarkstart and bookmarkend instance.
                    int indexOfText = textRange.OwnerParagraph.Items.IndexOf(textRange);
                    BookmarkStart bookmarkStart = new BookmarkStart(document, "bkmk1");
                    BookmarkEnd bookmarkEnd = new BookmarkEnd(document, "bkmk1");
                    //Add bookmarkstart before selected text.
                    textRange.OwnerParagraph.Items.Insert(indexOfText, bookmarkStart);
                    //Add bookmarkend after selected text
                    textRange.OwnerParagraph.Items.Insert(indexOfText + 2, bookmarkEnd);
                    textSelection = document.Find("The table below lists the main characteristics the giant panda shares with bears and red pandas.", false, true);
                    //Gets the found text as single text range
                    textRange = textSelection.GetAsOneRange();
                    //Add bookmark to the selected text range.
                    //Create bookmarkstart and bookmarkend instance.
                    indexOfText = textRange.OwnerParagraph.Items.IndexOf(textRange);
                    bookmarkStart = new BookmarkStart(document, "bkmk2");
                    bookmarkEnd = new BookmarkEnd(document, "bkmk2");
                    //Add bookmarkstart before selected text.
                    textRange.OwnerParagraph.Items.Insert(indexOfText, bookmarkStart);
                    //Add bookmarkend after selected text
                    textRange.OwnerParagraph.Items.Insert(indexOfText + 2, bookmarkEnd);
                    textSelection = document.Find("Did you know that the giant panda may actually be a raccoon", false, true);
                    //Gets the found text as single text range
                    textRange = textSelection.GetAsOneRange();
                    //Add bookmark to the selected text range.
                    //Create bookmarkstart and bookmarkend instance.
                    indexOfText = textRange.OwnerParagraph.Items.IndexOf(textRange);
                    bookmarkStart = new BookmarkStart(document, "bkmk3");
                    bookmarkEnd = new BookmarkEnd(document, "bkmk3");
                    //Add bookmarkstart before selected text.
                    textRange.OwnerParagraph.Items.Insert(indexOfText, bookmarkStart);
                    //Add bookmarkend after selected text
                    textRange.OwnerParagraph.Items.Insert(indexOfText + 2, bookmarkEnd);

                    //Get all bookmarks from Word document using FindAllItemsByProperty
                    //Find all bkmarkStart by EntityType in Word document.
                    List<Entity> bkmarkStarts = document.FindAllItemsByProperty(EntityType.BookmarkStart, null, null);
                    //Create an list Bookmarks of type string
                    List<string> BookmarksContent = new List<string>();
                    //Iterate bookmarkCollection to get the bookmark content.
                    foreach (Entity bkmarkStart in bkmarkStarts)
                    {
                        BookmarkStart book = bkmarkStart as BookmarkStart;
                        //Get the bookmark name
                        string name = book.Name;
                        //Creates the bookmark navigator instance to access the bookmark
                        BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(document);
                        //Moves the virtual cursor to the location before the end of the bookmark
                        bookmarkNavigator.MoveToBookmark(name);
                        //Gets the bookmark content as worddocument
                        WordDocumentPart part = bookmarkNavigator.GetContent();
                        WordDocument tempDoc = part.GetAsWordDocument();
                        //Get the bookmark content from the document.
                        string text = tempDoc.GetText();
                        //Adds the bookmark content into the list
                        BookmarksContent.Add(text);
                        Console.WriteLine("Bookmark content: ");
                        Console.WriteLine(text);
                        tempDoc.Close();
                        tempDoc.Dispose();
                        part.Close();

                    }
                    Console.ReadLine();
                    //Creates file stream.
                    using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}