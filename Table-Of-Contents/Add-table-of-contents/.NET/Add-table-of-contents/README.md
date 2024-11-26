# Add and update Table of Contents in Word document using C#

The Syncfusion [.NET Word Library](https://www.syncfusion.com/document-processing/word-framework/net/word-library) (DocIO) enables you to create, read, and edit Word documents programmatically without Microsoft Word or interop dependencies. Using this library, you can **add and update a Table of Contents (TOC) in a Word document** using C#.

## Steps to add and update Table of Contents (TOC) programmatically

Step 1: Create a new .NET Core console application project.

Step 2: Install the [Syncfusion.DocIORenderer.Net.Core](https://www.nuget.org/packages/Syncfusion.DocIORenderer.Net.Core) NuGet package as a reference to your project from [NuGet.org](https://www.nuget.org/).

Step 3: Include the following namespaces in the Program.cs file.

```csharp
using Syncfusion.DocIO; 
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using System.IO;
```

Step 4: Add the following code snippet in Program.cs file to add and update TOC in the Word document.

```csharp
//Creates a new Word document.
using (WordDocument document = new WordDocument())
{
    //Adds the section into the Word document.
    IWSection section = document.AddSection();
    string paraText = "AdventureWorks Cycles, the fictitious company on which the AdventureWorks sample databases are based, is a large, multinational manufacturing company.";
    //Adds the paragraph into the created section.
    IWParagraph paragraph = section.AddParagraph();
    //Appends the TOC field with LowerHeadingLevel and UpperHeadingLevel to determines the TOC entries.
    paragraph.AppendTOC(1, 3);
    //Adds the section into the Word document.
    section = document.AddSection();
    //Adds the paragraph into the created section.
    paragraph = section.AddParagraph();
    //Adds the text for the headings.
    paragraph.AppendText("First Chapter");
    //Sets a built-in heading style.
    paragraph.ApplyStyle(BuiltinStyle.Heading1);
    //Adds the text into the paragraph.
    section.AddParagraph().AppendText(paraText);
    //Adds the section into the Word document.
    section = document.AddSection();
    //Adds the paragraph into the created section.
    paragraph = section.AddParagraph();
    //Adds the text for the headings.
    paragraph.AppendText("Second Chapter");
    //Sets a built-in heading style.
    paragraph.ApplyStyle(BuiltinStyle.Heading2);
    //Adds the text into the paragraph.
    section.AddParagraph().AppendText(paraText);
    //Adds the section into the Word document.
    section = document.AddSection();
    //Adds the paragraph into the created section
    paragraph = section.AddParagraph();
    //Adds the text into the headings.
    paragraph.AppendText("Third Chapter");
    //Sets a built-in heading style.
    paragraph.ApplyStyle(BuiltinStyle.Heading3);
    //Adds the text into the paragraph.
    section.AddParagraph().AppendText(paraText);
    //Updates the table of contents.
    document.UpdateTableOfContents();
    //Creates file stream.
    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
    {
        //Saves the Word document to file stream.
        document.Save(outputFileStream, FormatType.Docx);
    }
}
```

More information about the Table of Contents can be found in this [documentation](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-table-of-contents) section.