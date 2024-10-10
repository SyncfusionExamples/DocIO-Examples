using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Replace_content_with_bookmark
{
    internal class Program
    {
        static void Main(string[] args)
        {
            FileStream inputOne = new FileStream(Path.GetFullPath("Data/SourceDocument.docx"), FileMode.Open, FileAccess.Read);
            WordDocument inputOneDoc = new WordDocument(inputOne, FormatType.Docx);
            FileStream inputTwo = new FileStream(Path.GetFullPath(@"Data/DestinationDocument.docx"), FileMode.Open, FileAccess.Read);
            WordDocument inputTwoDoc = new WordDocument(inputTwo, FormatType.Docx);
            DocxReplaceTextWithDocPart(inputOneDoc, inputTwoDoc, "Text one", "bkmk1");
            DocxReplaceTextWithDocPart(inputOneDoc, inputTwoDoc, "Text two", "bkmk2");
            FileStream output = new FileStream(Path.GetFullPath("Output/Output.docx"), FileMode.Create, FileAccess.Write);
            inputOneDoc.Save(output, FormatType.Docx);
            inputOneDoc.Close();
            inputTwoDoc.Close();
            inputOne.Close();
            inputTwo.Close();
            output.Close();
        }

        private static void DocxReplaceTextWithDocPart(WordDocument document, WordDocument sourceDoc, string tokenToFind, string textBookmark)
        {
            string bookmarkRef = textBookmark + "_bm";

            //Find the start token.
            TextSelection start = document.Find(tokenToFind, true, true);
            if (start != null)
            {
                WTextRange startText = start.GetAsOneRange();
                WParagraph startParagraph = startText.OwnerParagraph;
                //Get the item index of the start token in the paragraph
                int index = startParagraph.Items.IndexOf(startText);
                //Remove the start token at the specified index
                startParagraph.Items.Remove(startText);
                //Create and insert a Bookmark start at the index of the start token
                BookmarkStart bookmarkStart = new BookmarkStart(document, bookmarkRef);
                startParagraph.Items.Insert(index, bookmarkStart);
                startParagraph.AppendBookmarkEnd(bookmarkRef);

                //Open the document that contains the text to replace
                //For instance, the document contains Bookmark named "DocIO" and the contents of that 
                //bookmark should replace the content in above document 
                //Creates the bookmark navigator instance to access the bookmark
                if (sourceDoc.Bookmarks.FindByName(textBookmark) != null)
                {
                    BookmarksNavigator bookmarksNavigator = new BookmarksNavigator(sourceDoc);
                    //Moves the virtual cursor to the location before the end of the bookmark "DocIO"
                    bookmarksNavigator.MoveToBookmark(textBookmark);
                    //Gets the bookmark content
                    WordDocumentPart wordDocumentPart = bookmarksNavigator.GetContent();
                    //Creates the bookmark navigator instance to access the bookmark
                    bookmarksNavigator = new BookmarksNavigator(document);
                    //Moves the virtual cursor to the location before the end of the bookmark "Bookmark"
                    bookmarksNavigator.MoveToBookmark(bookmarkRef);

                    //Get the destination para before replacing
                    WParagraph destinationPara = bookmarksNavigator.CurrentBookmark.BookmarkStart.OwnerParagraph;
                    //Get the list style name
                    string listStyleName = destinationPara.ListFormat.CustomStyleName;
                    //Get the first line indent value
                    float firstLineIndent = destinationPara.ParagraphFormat.FirstLineIndent;
                    //Get the left indent value
                    float leftIndent = destinationPara.ParagraphFormat.LeftIndent;

                    //Replace the selected text with another Word document content
                    bookmarksNavigator.ReplaceContent(wordDocumentPart);
                    //Apply the list style, first line indent and left indent values after replacing
                    destinationPara.ListFormat.ApplyStyle(listStyleName);
                    destinationPara.ParagraphFormat.FirstLineIndent = firstLineIndent;
                    destinationPara.ParagraphFormat.LeftIndent = leftIndent;
                }
                else
                {
                    //Creates the bookmark navigator instance to access the bookmark
                    BookmarksNavigator bookmarksNavigator = new BookmarksNavigator(document);
                    //Moves the virtual cursor to the location before the end of the bookmark "Bookmark"
                    bookmarksNavigator.MoveToBookmark(bookmarkRef);
                    //Replace the selected text with another Word document content
                    bookmarksNavigator.ReplaceBookmarkContent(string.Empty, true);
                }
            }
        }
    }
}
