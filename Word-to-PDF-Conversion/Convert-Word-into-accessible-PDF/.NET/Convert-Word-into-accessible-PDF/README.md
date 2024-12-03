# Convert Word document to PDF/UA using C#

The Syncfusion&reg; [.NET Word Library](https://www.syncfusion.com/document-processing/word-framework/net/word-library) (DocIO) enables you to create, read, edit, and convert Word documents programmatically without Microsoft Word or interop dependencies. Using this library, you can **convert a Word document to PDF/UA** using C#.

## Steps to convert Word to PDF/UA programmatically

Step 1: Create a new .NET Core console application project.

Step 2: Install the [Syncfusion.DocIORenderer.Net.Core](https://www.nuget.org/packages/Syncfusion.DocIORenderer.Net.Core) NuGet package as a reference to your project from [NuGet.org](https://www.nuget.org/).

Step 3: Include the following namespaces in the Program.cs file.

```csharp
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using System.IO;
```

Step 4: Add the following code snippet in Program.cs file to convert a Word document to PDF/UA.

```csharp
using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open))
{
    //Loads an existing Word document.
    using (WordDocument wordDocument = new WordDocument(fileStream, Syncfusion.DocIO.FormatType.Automatic))
    {
        //Creates an instance of DocIORenderer.
        using (DocIORenderer renderer = new DocIORenderer())
        {
            //Sets true to preserve document structured tags in the converted PDF document.
            renderer.Settings.AutoTag = true;
            //Converts Word document into PDF document.
            using (PdfDocument pdfDocument = renderer.ConvertToPDF(wordDocument))
            {
                //Saves the PDF file to file system.    
                using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Output.pdf"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
                {
                    pdfDocument.Save(outputStream);
                }
            }
        }
    }
}
```

More information about Word to PDF/UA can be found in this [documentation](https://help.syncfusion.com/document-processing/word/conversions/word-to-pdf/net/word-to-pdf#accessible-pdf-document) section.