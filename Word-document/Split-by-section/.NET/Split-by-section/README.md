# Split Word document by Section using C#

The Syncfusion [.NET Word Library](https://www.syncfusion.com/document-processing/word-framework/net/word-library) (DocIO) enables you to create, read, and edit Word documents programmatically without Microsoft Word or interop dependencies. Using this library, you can **split a Word document by sections** using C#.

## Steps to split Word document by sections programmatically

Step 1: Create a new .NET Core console application project.

Step 2: Install the [Syncfusion.DocIO.Net.Core](https://www.nuget.org/packages/Syncfusion.DocIO.Net.Core) NuGet package as a reference to your project from [NuGet.org](https://www.nuget.org/).

Step 3: Include the following namespaces in the Program.cs file.

```csharp
using Syncfusion.DocIO; 
using Syncfusion.DocIO.DLS;
using System.IO;
```

Step 4: Add the following code snippet in Program.cs file to split Word document by sections.

```csharp
using (FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    //Load the template document as stream
    using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
    {
        int fileId = 1;
        //Iterate each section from Word document
        foreach (WSection section in document.Sections)
        {
            //Create new Word document
            using (WordDocument newDocument = new WordDocument())
            {
                //Add cloned section into new Word document
                newDocument.Sections.Add(section.Clone());
                //Saves the Word document to  MemoryStream
                using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Section") + fileId + ".docx", FileMode.OpenOrCreate, FileAccess.ReadWrite))
                {
                    newDocument.Save(outputStream, FormatType.Docx);
                }
            }
            fileId++;
        }
    }
}
```

More information about splitting a Word document can be found in this [documentation](https://help.syncfusion.com/document-processing/word/word-library/net/word-document/split-word-documents) section.