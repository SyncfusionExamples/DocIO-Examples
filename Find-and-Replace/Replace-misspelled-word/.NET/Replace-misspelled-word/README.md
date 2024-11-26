# Find and replace text in Word document using C#

The Syncfusion [.NET Word Library](https://www.syncfusion.com/document-processing/word-framework/net/word-library) (DocIO) enables you to create, read, and edit Word documents programmatically without Microsoft Word or interop dependencies. Using this library, you can **find and replace text in a Word document** using C#.

## Steps to find and replace text programmatically

Step 1: Create a new .NET Core console application project.

Step 2: Install the [Syncfusion.DocIO.Net.Core](https://www.nuget.org/packages/Syncfusion.DocIO.Net.Core) NuGet package as a reference to your project from [NuGet.org](https://www.nuget.org/).

Step 3: Include the following namespaces in the Program.cs file.

```csharp
using Syncfusion.DocIO; 
using Syncfusion.DocIO.DLS;
using System.IO;
```

Step 4: Add the following code snippet in Program.cs file to find and replace text in the Word document.

```csharp
using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    //Opens an existing Word document.
    using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
    {
        //Finds all occurrences of a misspelled word and replaces with properly spelled word.
        document.Replace("Cyles", "Cycles", true, true);
        //Creates file stream.
        using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
        {
            //Saves the Word document to file stream.
            document.Save(outputFileStream, FormatType.Docx);
        }
    }
}
```

More information about find and replace functionality can be found in this [documentation](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-find-and-replace) section.
