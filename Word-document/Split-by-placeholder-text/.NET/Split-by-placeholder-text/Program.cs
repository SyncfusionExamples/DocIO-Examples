using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace Split_a_document_by_placeholder_text
{
    class Program
    {
        static void Main(string[] args)
        {
            //Load an existing Word document into DocIO instance.
            FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
            {
                
                //Finds all the placeholder text in the Word document.
                TextSelection[] textSelections = document.FindAll(new Regex("<<(.*)>>"));
                if (textSelections != null)
                {
                    #region Insert bookmarks at placeholders
                    //Unique ID for each bookmark.
                    int bkmkId = 1;
                    //Collection to hold the inserted bookmarks.
                    List<string> bookmarks = new List<string>();
                    //Iterate each text selection.
                    for (int i = 0; i < textSelections.Length; i++)
                    {
                        #region Insert bookmark start before the placeholder
                        //Get the placeholder as WTextRange.
                        WTextRange textRange = textSelections[i].GetAsOneRange();
                        //Get the index of the placeholder text. 
                        WParagraph startParagraph = textRange.OwnerParagraph;
                        int index = startParagraph.ChildEntities.IndexOf(textRange);
                        string bookmarkName = "Bookmark_" + bkmkId;
                        //Add new bookmark to bookmarks collection.
                        bookmarks.Add(bookmarkName);
                        //Create bookmark start.
                        BookmarkStart bkmkStart = new BookmarkStart(document, bookmarkName);
                        //Insert the bookmark start before the start placeholder.
                        startParagraph.ChildEntities.Insert(index, bkmkStart);
                        //Remove the placeholder text.
                        textRange.Text = string.Empty;
                        #endregion

                        #region Insert bookmark end after the placeholder
                        i++;
                        //Get the placeholder as WTextRange.
                        textRange = textSelections[i].GetAsOneRange();
                        //Get the index of the placeholder text. 
                        WParagraph endParagraph = textRange.OwnerParagraph;
                        index = endParagraph.ChildEntities.IndexOf(textRange);
                        //Create bookmark end.
                        BookmarkEnd bkmkEnd = new BookmarkEnd(document, bookmarkName);
                        //Insert the bookmark end after the end placeholder.
                        endParagraph.ChildEntities.Insert(index + 1, bkmkEnd);
                        bkmkId++;
                        //Remove the placeholder text.
                        textRange.Text = string.Empty;
                        #endregion

                    }
                    #endregion
                    #region Split bookmark content into separate documents 
                    BookmarksNavigator bookmarksNavigator = new BookmarksNavigator(document);
                    int fileIndex = 1;
                    foreach (string bookmark in bookmarks)
                    {
                        //Move the virtual cursor to the location before the end of the bookmark.
                        bookmarksNavigator.MoveToBookmark(bookmark);
                        //Get the bookmark content as WordDocumentPart.
                        WordDocumentPart wordDocumentPart = bookmarksNavigator.GetContent();
                        //Save the WordDocumentPart as separate Word document.
                        using (WordDocument newDocument = wordDocumentPart.GetAsWordDocument())
                        {
                            //Save the Word document to file stream.
                            using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Placeholder_" + fileIndex + ".docx"), FileMode.Create, FileAccess.ReadWrite))
                            {
                                newDocument.Save(outputFileStream, FormatType.Docx);
                            }
                        }
                        fileIndex++;
                    }
                    #endregion
                }
            }
        }
    }
}
