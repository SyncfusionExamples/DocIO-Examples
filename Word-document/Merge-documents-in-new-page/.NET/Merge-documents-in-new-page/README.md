# Merge Word documents using C#

The Syncfusion [.NET Word Library](https://www.syncfusion.com/document-processing/word-framework/net/word-library) (DocIO) enables you to create, read, and edit Word documents programmatically without Microsoft Word or interop dependencies. Using this library, you can **merge Word documents** using C#.

## Steps to merge Word documents programmatically

Step 1: Create a new .NET Core console application project.

Step 2: Install the [Syncfusion.DocIO.Net.Core](https://www.nuget.org/packages/Syncfusion.DocIO.Net.Core) NuGet package as a reference to your project from [NuGet.org](https://www.nuget.org/).

Step 3: Include the following namespaces in the Program.cs file.

```csharp
using Syncfusion.DocIO; 
using Syncfusion.DocIO.DLS;
using System.IO; 
```

Step 4: Add the following code snippet in Program.cs file to merge Word documents.

```csharp
using (FileStream sourceStreamPath = new FileStream(Path.GetFullPath(@"Data/SourceDocument.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    //Opens an source document from file system through constructor of WordDocument class.
    using (WordDocument sourceDocument = new WordDocument(sourceStreamPath, FormatType.Automatic))
    {
        using (FileStream destinationStreamPath = new FileStream(Path.GetFullPath(@"Data/DestinationDocument.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        {
            //Opens the destination document.
            using (WordDocument destinationDocument = new WordDocument(destinationStreamPath, FormatType.Automatic))
            {
                //Imports the contents of source document at the end of destination document.
                destinationDocument.ImportContent(sourceDocument, ImportOptions.UseDestinationStyles);
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    destinationDocument.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
```

More information about merging Word documents can be found in this [documentation](https://help.syncfusion.com/document-processing/word/word-library/net/word-document/merging-word-documents) section.