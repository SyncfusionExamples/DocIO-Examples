# Convert Word document to Image using C#

The Syncfusion [.NET Word Library](https://www.syncfusion.com/document-processing/word-framework/net/word-library) (DocIO) enables you to create, read, edit, and convert Word documents programmatically without Microsoft Word or interop dependencies. Using this library, you can **convert a Word document to Image** using C#.

## Steps to convert Word to Image programmatically

Step 1: Create a new .NET Core console application project.

Step 2: Install the [Syncfusion.DocIORenderer.Net.Core](https://www.nuget.org/packages/Syncfusion.DocIORenderer.Net.Core) NuGet package as a reference to your project from [NuGet.org](https://www.nuget.org/).

Step 3: Include the following namespaces in the Program.cs file.

```csharp
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using System.IO;
```

Step 4: Add the following code snippet in Program.cs file to convert a Word document to image.

```csharp
using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open))
{
    //Loads an existing Word document.
    using (WordDocument wordDocument = new WordDocument(fileStream, FormatType.Automatic))
    {
        //Creates an instance of DocIORenderer.
        using (DocIORenderer renderer = new DocIORenderer())
        {
            //Convert the first page of the Word document into an image.
            Stream imageStream = wordDocument.RenderAsImages(0, ExportImageFormat.Jpeg);
            //Resets the stream position.
            imageStream.Position = 0;
            //Creates the output image file stream.
            using (FileStream fileStreamOutput = File.Create(Path.GetFullPath(@"Output/Output.jpeg")))
            {
                //Copies the converted image stream into created output stream.
                imageStream.CopyTo(fileStreamOutput);
            }
        }
    }
}
```

More information about Word to Image conversion can be found in this [documentation](https://help.syncfusion.com/document-processing/word/conversions/word-to-image/net/word-to-image) section.