# Split Word document by Bookmarks using C#

The Syncfusion&reg; [.NET Word Library](https://www.syncfusion.com/document-processing/word-framework/net/word-library) (DocIO) empowers you to create, read, and edit Word documents programmatically without Microsoft Word or interop dependencies. Using this library, you can **split a Word document by bookmarks** using C#.

## Steps to split Word document by bookmarks programmatically

Step 1: Create a new .NET Core console application project.

Step 2: Install the [Syncfusion.DocIO.Net.Core](https://www.nuget.org/packages/Syncfusion.DocIO.Net.Core) NuGet package as a reference to your project from [NuGet.org](https://www.nuget.org/).

Step 3: Include the following namespaces in the Program.cs file.

```csharp
using Syncfusion.DocIO; 
using Syncfusion.DocIO.DLS;
using System.IO;
```

Step 4: Add the following code snippet in Program.cs file to split Word document by bookmarks.

```csharp
//Load an existing Word document.
FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
{
    //Create the bookmark navigator instance to access the bookmark.
    BookmarksNavigator bookmarksNavigator = new BookmarksNavigator(document);
    BookmarkCollection bookmarkCollection = document.Bookmarks;
    //Iterate each bookmark in Word document.
    foreach (Bookmark bookmark in bookmarkCollection)
    {
        //Move the virtual cursor to the location before the end of the bookmark.
        bookmarksNavigator.MoveToBookmark(bookmark.Name);
        //Get the bookmark content as WordDocumentPart.
        WordDocumentPart documentPart = bookmarksNavigator.GetContent();
        //Save the WordDocumentPart as separate Word document
        using (WordDocument newDocument = documentPart.GetAsWordDocument())
        {
            //Save the Word document to file stream.
            using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/" + bookmark.Name + ".docx"), FileMode.Create, FileAccess.ReadWrite))
            {
                newDocument.Save(outputFileStream, FormatType.Docx);
            }
        }
    }
} 
```

More information about splitting a Word document can be found in this [documentation](https://help.syncfusion.com/document-processing/word/word-library/net/word-document/split-word-documents) section.