# Replace bookmark content in a Word document using C#

The Syncfusion&reg; [.NET Word Library](https://www.syncfusion.com/document-processing/word-framework/net/word-library) (DocIO) enables you to create, read, and edit Word documents programmatically without Microsoft Word or interop dependencies. Using this library, you can **replace bookmark content in a Word document** using C#.

## Steps to replace bookmark content programmatically

Step 1: Create a new .NET Core console application project.

Step 2: Install the [Syncfusion.DocIO.Net.Core](https://www.nuget.org/packages/Syncfusion.DocIO.Net.Core) NuGet package as a reference to your project from [NuGet.org](https://www.nuget.org/).

Step 3: Include the following namespaces in the Program.cs file.

```csharp
using Syncfusion.DocIO; 
using Syncfusion.DocIO.DLS;
using System.IO;
```

Step 4: Add the following code snippet in Program.cs file to replace bookmark content in the Word document.

```csharp
using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    //Opens an existing Word document.
    using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
    {
        //Creates the bookmark navigator instance to access the bookmark.
        BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(document);
        //Moves the virtual cursor to the location before the end of the bookmark "Northwind".
        bookmarkNavigator.MoveToBookmark("Northwind");
        //Gets the bookmark content.
        TextBodyPart textBodyPart = bookmarkNavigator.GetBookmarkContent();
        document.AddSection();
        IWParagraph paragraph = document.LastSection.AddParagraph();
        paragraph.AppendText("Northwind Database is a set of tables containing data fitted into predefined categories.");
        //Adds the new bookmark into Word document.
        paragraph.AppendBookmarkStart("bookmark_empty");
        paragraph.AppendBookmarkEnd("bookmark_empty");
        //Moves the virtual cursor to the location before the end of the bookmark "bookmark_empty".
        bookmarkNavigator.MoveToBookmark("bookmark_empty");
        //Replaces the bookmark content with text body part.
        bookmarkNavigator.ReplaceBookmarkContent(textBodyPart);
        //Creates file stream.
        using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite))
        {
            //Saves the Word document to file stream.
            document.Save(outputFileStream, FormatType.Docx);
        }
    }
}
```

More information about the bookmarks can be found in this [documentation](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-bookmarks) section.