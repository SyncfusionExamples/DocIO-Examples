using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;
using System.Text.RegularExpressions;

namespace Replace_text_inside_tag
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/DestinationDocument.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Open the destination Word document.
                using (WordDocument destinationDocument = new WordDocument(fileStreamPath, FormatType.Docx)) 
                {
                    using (FileStream sourceFileStream = new FileStream(Path.GetFullPath(@"Data/SourceDocument.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        //Open the source Word document.
                        using (WordDocument sourceDocument = new WordDocument(sourceFileStream, FormatType.Docx))
                        {
                            //Get the content between the tags in the source document as body part.
                            TextSelection[] textSelections = sourceDocument.FindSingleLine(new Regex("<SourceTag>(.*)</SourceTag>"));
                            if (textSelections != null)
                            {
                                TextBodyPart bodyPart = new TextBodyPart(destinationDocument);
                                for (int i = 1; i < textSelections.Length - 1; i++)
                                {
                                    WParagraph paragraph = new WParagraph(destinationDocument);
                                    foreach (var range in textSelections[i].GetRanges())
                                    {
                                        WTextRange textrange = range.Clone() as WTextRange;
                                        paragraph.ChildEntities.Add(textrange);
                                    }
                                    bodyPart.BodyItems.Add(paragraph);
                                }
                                //Replace the text between specified tags in the destination document using a bookmark 
                                //with the content from the source document.
                                ReplaceTextBetweenTags(destinationDocument, bodyPart);
                            }
                            //Create file stream.
                            using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite))
                            {
                                //Save the Word document to file stream.
                                destinationDocument.Save(outputFileStream, FormatType.Docx);
                            }
                        }
                    }
                }
            }
        }
        #region Helper Methods
        /// <summary>
        /// Replaces the content between specified start and end tags within the destination document using a bookmark. 
        /// </summary>
        /// <param name="destinationDocument">The Word document where the replacement is to be performed.</param>
        /// <param name="bodyPart">The content to insert between the specified start and end tags.</param>
        private static void ReplaceTextBetweenTags(WordDocument destinationDocument, TextBodyPart bodyPart)
        {
            //Define the start and end tags to identify the content to be replaced.
            string startTag = "<DestTag>";
            string endTag = "</DestTag>";
            //Create bookmark start and bookmark end.
            BookmarkStart bookmarkStart = new BookmarkStart(destinationDocument, "Adventure_Bkmk");
            BookmarkEnd bookmarkEnd = new BookmarkEnd(destinationDocument, "Adventure_Bkmk");

            //Find the start tag in the destination document.
            TextSelection textSelection = destinationDocument.Find(startTag, false, false);
            if (textSelection == null) return; //Exit if start tag is not found.

            //Add a bookmark start after the start tag location.
            WTextRange startTagTextRange = textSelection.GetAsOneRange();
            WParagraph startTagParagraph = startTagTextRange.OwnerParagraph;
            int startTagIndex = startTagParagraph.ChildEntities.IndexOf(startTagTextRange);
            startTagParagraph.Items.Insert(startTagIndex + 1, bookmarkStart);

            //Find the end tag in the destination document.
            textSelection = destinationDocument.Find(endTag, false, false);
            if (textSelection == null) return; // Exit if end tag is not found

            //Add a bookmark end at the end tag location (before end tag).
            WTextRange endTagTextRange = textSelection.GetAsOneRange();
            WParagraph endTagParagraph = endTagTextRange.OwnerParagraph;
            int endTagIndex = endTagParagraph.ChildEntities.IndexOf(endTagTextRange);
            endTagParagraph.Items.Insert(endTagIndex, bookmarkEnd);

            //Create the bookmark navigator instance to access the bookmark.
            BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(destinationDocument);
            //Move the virtual cursor to the location of the bookmark "Adventure_Bkmk".
            bookmarkNavigator.MoveToBookmark("Adventure_Bkmk");
            //Replace the bookmark content with body part.
            bookmarkNavigator.ReplaceBookmarkContent(bodyPart);

            //Remove the bookmark from the destination document after replacing the content.
            Bookmark bookmark = destinationDocument.Bookmarks.FindByName("Adventure_Bkmk");
            if (bookmark != null)
                destinationDocument.Bookmarks.Remove(bookmark);
        }
        #endregion
    }
}