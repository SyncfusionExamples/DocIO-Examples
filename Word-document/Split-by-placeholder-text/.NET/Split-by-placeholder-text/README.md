# Split Word document by PlaceHolder using C#

The Syncfusion&reg; [.NET Word Library](https://www.syncfusion.com/document-processing/word-framework/net/word-library) (DocIO) allows you to create, read, and edit Word documents programmatically without Microsoft Word or interop dependencies. Using this library, you can **split a Word document by placeholders** using C#.

## Steps to split Word document by placeholders programmatically

Step 1: Create a new .NET Core console application project.

Step 2: Install the [Syncfusion.DocIO.Net.Core](https://www.nuget.org/packages/Syncfusion.DocIO.Net.Core) NuGet package as a reference to your project from [NuGet.org](https://www.nuget.org/).

Step 3: Include the following namespaces in the Program.cs file.

```csharp
using Syncfusion.DocIO; 
using Syncfusion.DocIO.DLS;
using System.IO;
using System.Collections.Generic;
using System.Text.RegularExpressions;
```

Step 4: Add the following code snippet in Program.cs file to split Word document by placeholders.

```csharp
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
```

More information about the mail merge can be found in this [documentation](https://help.syncfusion.com/file-formats/docio/working-with-mailmerge) section.