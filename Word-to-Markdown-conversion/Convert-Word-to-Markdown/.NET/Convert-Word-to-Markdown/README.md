# Convert Word document to Markdown using C#

The Syncfusion&reg; [.NET Word Library](https://www.syncfusion.com/document-processing/word-framework/net/word-library) (DocIO) enables you to create, read, edit, and convert Word documents programmatically without Microsoft Word or interop dependencies. Using this library, you can **convert a Word document to Markdown** using C#.

## Steps to convert Word to Markdown programmatically

Step 1: Create a new .NET Core console application project.

Step 2: Install the [Syncfusion.DocIO.Net.Core](https://www.nuget.org/packages/Syncfusion.DocIO.Net.Core) NuGet package as a reference to your project from [NuGet.org](https://www.nuget.org/).

Step 3: Include the following namespaces in the Program.cs file.

```csharp
using Syncfusion.DocIO; 
using Syncfusion.DocIO.DLS;
using System.IO;
```

Step 4: Add the following code snippet in Program.cs file to convert a Word document to Markdown.

```csharp
//Open a file as a stream.
using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    //Load the file stream into a Word document.
    using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
    {
        //Create a file stream.
        using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output.md"), FileMode.Create, FileAccess.ReadWrite))
        {
            //Save a Markdown file to the file stream.
            document.Save(outputFileStream, FormatType.Markdown);
        }
    }
}
```

More information about Word to Markdown conversion can be found in this [documentation](https://help.syncfusion.com/document-processing/word/word-library/net/convert-word-document-to-markdown-in-csharp) section.