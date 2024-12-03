# Convert Word document to PDF/A using C#

The Syncfusion&reg; [.NET Word Library](https://www.syncfusion.com/document-processing/word-framework/net/word-library) (DocIO) enables you to create, read, edit, and convert Word documents programmatically without Microsoft Word or interop dependencies. Using this library, you can **convert a Word document to PDF/A** using C#.

## Steps to convert Word to PDF/A programmatically

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

Step 4: Add the following code snippet in Program.cs file to convert a Word document to PDF/A.

```csharp
using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open))
{
    //Loads an existing Word document.
    using (WordDocument wordDocument = new WordDocument(fileStream, Syncfusion.DocIO.FormatType.Automatic))
    {
        //Creates an instance of DocIORenderer.
        using (DocIORenderer renderer = new DocIORenderer())
        {
            //Set the conformance for PDF/A-1b conversion.
            renderer.Settings.PdfConformanceLevel = PdfConformanceLevel.Pdf_A1B;
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

More information about converting Word document to PDF with conformance can be found in this [documentation](https://help.syncfusion.com/document-processing/word/conversions/word-to-pdf/net/word-to-pdf#pdf-conformance-level) section.