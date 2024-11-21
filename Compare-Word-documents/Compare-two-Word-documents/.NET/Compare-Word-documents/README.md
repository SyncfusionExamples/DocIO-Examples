# Compare Word documents using C#

The Syncfusion [.NET Word Library](https://www.syncfusion.com/document-processing/word-framework/net/word-library) (DocIO) enables you to create, read, and edit Word documents programmatically without Microsoft Word or interop dependencies. Using this library, you can **compare Word documents** using C#.

## Steps to compare Word documents programmatically

Step 1: Create a new .NET Core console application project.

Step 2: Install the [Syncfusion.DocIO.Net.Core](https://www.nuget.org/packages/Syncfusion.DocIO.Net.Core) NuGet package as a reference to your project from [NuGet.org](https://www.nuget.org/).

Step 3: Include the following namespaces in the Program.cs file.

```csharp
using Syncfusion.DocIO; 
using Syncfusion.DocIO.DLS;
```

Step 4: Add the following code snippet in Program.cs file to compare Word documents.

```csharp
//Loads the original document.
using (FileStream originalDocumentStreamPath = new FileStream(Path.GetFullPath(@"Data/OriginalDocument.docx"), FileMode.Open, FileAccess.Read))
{
    using (WordDocument originalDocument = new WordDocument(originalDocumentStreamPath, FormatType.Docx))
    {
        //Loads the revised document
        using (FileStream revisedDocumentStreamPath = new FileStream(Path.GetFullPath(@"Data/RevisedDocument.docx"), FileMode.Open, FileAccess.Read))
        {
            using (WordDocument revisedDocument = new WordDocument(revisedDocumentStreamPath, FormatType.Docx))
            {
                //Compare the original and revised Word documents.
                originalDocument.Compare(revisedDocument);
                //Save the Word document.
                using (FileStream fileStreamOutput = File.Create(Path.GetFullPath("Output/Output.docx")))
                {
                    originalDocument.Save(fileStreamOutput, FormatType.Docx);
                }
            }
        }     
    }               
}
```

More information about comparing Word document can be found in this [documentation](https://help.syncfusion.com/document-processing/word/word-library/net/word-document/compare-word-documents) section.