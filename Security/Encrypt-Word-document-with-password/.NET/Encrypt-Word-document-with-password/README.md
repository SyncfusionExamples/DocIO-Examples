# Encrypt Word document using C#

The Syncfusion&reg; [.NET Word Library](https://www.syncfusion.com/document-processing/word-framework/net/word-library) (DocIO) enables you to create, read, and edit Word documents programmatically without Microsoft Word or interop dependencies. Using this library, you can **encrypt Word documents** using C#.

## Steps to encrypt a Word document programmatically

Step 1: Create a new .NET Core console application project.

Step 2: Install the [Syncfusion.DocIO.Net.Core](https://www.nuget.org/packages/Syncfusion.DocIO.Net.Core) NuGet package as a reference to your project from [NuGet.org](https://www.nuget.org/).

Step 3: Include the following namespaces in the Program.cs file.

```csharp
using Syncfusion.DocIO; 
using Syncfusion.DocIO.DLS;
using System.IO;
```

Step 4: Add the following code snippet in Program.cs file to encrypt a Word document.

```csharp
using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.ReadWrite))
{
    //Opens the template document.
    using (WordDocument document = new WordDocument(fileStream, FormatType.Docx))
    {
        //Encrypts the Word document with a password.
        document.EncryptDocument("syncfusion");
        //Creates file stream.
        using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
        {
            //Saves the Word document to file stream.
            document.Save(outputStream, FormatType.Docx);
        }
    }
}
```

More information about the encrypt and decrypt options can be found in this [documentation](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-security) section.