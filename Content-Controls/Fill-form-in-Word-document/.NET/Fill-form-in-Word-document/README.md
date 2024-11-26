# Form filling in Word document using C#

The Syncfusion [.NET Word Library](https://www.syncfusion.com/document-processing/word-framework/net/word-library) (DocIO) enables you to create, read, and edit Word documents programmatically without Microsoft Word or interop dependencies. Using this library, you can **fill forms in a Word document** using C#.

## Steps to fill forms programmatically

Step 1: Create a new .NET Core console application project.

Step 2: Install the [Syncfusion.DocIO.Net.Core](https://www.nuget.org/packages/Syncfusion.DocIO.Net.Core) NuGet package as a reference to your project from [NuGet.org](https://www.nuget.org/).

Step 3: Include the following namespaces in the Program.cs file.

```csharp
using Syncfusion.DocIO; 
using Syncfusion.DocIO.DLS;
using System.IO;
```

Step 4: Add the following code snippet in Program.cs file to fill forms in the Word document.

```csharp
using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    //Creates a new Word document.
    using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
    {
        IWSection sec = document.LastSection;
        InlineContentControl inlineCC;
        InlineContentControl dropDownCC;
        WTable table1 = sec.Tables[1] as WTable;
        WTableRow row1 = table1.Rows[1];

        #region General Information
        //Fill the name.
        WParagraph cellPara1 = row1.Cells[0].ChildEntities[1] as WParagraph;
        inlineCC = cellPara1.ChildEntities.LastItem as InlineContentControl;
        WTextRange text = new WTextRange(document);
        text.ApplyCharacterFormat(inlineCC.BreakCharacterFormat);
        text.Text = "Steve Jobs";
        inlineCC.ParagraphItems.Add(text);
        //Fill the date of birth.
        cellPara1 = row1.Cells[0].ChildEntities[3] as WParagraph;
        inlineCC = cellPara1.ChildEntities.LastItem as InlineContentControl;
        text = new WTextRange(document);
        text.ApplyCharacterFormat(inlineCC.BreakCharacterFormat);
        text.Text = "06/01/1994";
        inlineCC.ParagraphItems.Add(text);
        //Fill the address.
        cellPara1 = row1.Cells[0].ChildEntities[5] as WParagraph;
        inlineCC = cellPara1.ChildEntities.LastItem as InlineContentControl;
        text = new WTextRange(document);
        text.ApplyCharacterFormat(inlineCC.BreakCharacterFormat);
        text.Text = "2501 Aerial Center Parkway.";
        inlineCC.ParagraphItems.Add(text);
        text = new WTextRange(document);
        text.ApplyCharacterFormat(inlineCC.BreakCharacterFormat);
        text.Text = "Morrisville, NC 27560.";
        inlineCC.ParagraphItems.Add(text);
        text = new WTextRange(document);
        text.ApplyCharacterFormat(inlineCC.BreakCharacterFormat);
        text.Text = "USA.";
        inlineCC.ParagraphItems.Add(text);
        //Fill the phone no.
        cellPara1 = row1.Cells[0].ChildEntities[7] as WParagraph;
        inlineCC = cellPara1.ChildEntities.LastItem as InlineContentControl;
        text = new WTextRange(document);
        text.ApplyCharacterFormat(inlineCC.BreakCharacterFormat);
        text.Text = "+1 919.481.1974";
        inlineCC.ParagraphItems.Add(text);
        //Fill the email id.
        cellPara1 = row1.Cells[0].ChildEntities[9] as WParagraph;
        inlineCC = cellPara1.ChildEntities.LastItem as InlineContentControl;
        text = new WTextRange(document);
        text.ApplyCharacterFormat(inlineCC.BreakCharacterFormat);
        text.Text = "steve123@email.com";
        inlineCC.ParagraphItems.Add(text);
        #endregion

        #region Educational Information
        table1 = sec.Tables[2] as WTable;
        row1 = table1.Rows[1];
        //Fill the education type.
        cellPara1 = row1.Cells[0].ChildEntities[1] as WParagraph;
        dropDownCC = cellPara1.ChildEntities.LastItem as InlineContentControl;
        text = new WTextRange(document);
        text.ApplyCharacterFormat(dropDownCC.BreakCharacterFormat);
        text.Text = dropDownCC.ContentControlProperties.ContentControlListItems[1].DisplayText;
        dropDownCC.ParagraphItems.Add(text);
        //Fill the university.
        cellPara1 = row1.Cells[0].ChildEntities[3] as WParagraph;
        inlineCC = cellPara1.ChildEntities.LastItem as InlineContentControl;
        text = new WTextRange(document);
        text.ApplyCharacterFormat(dropDownCC.BreakCharacterFormat);
        text.Text = "Michigan University";
        inlineCC.ParagraphItems.Add(text);
        //Fill the C# experience level.
        cellPara1 = row1.Cells[0].ChildEntities[7] as WParagraph;
        dropDownCC = cellPara1.ChildEntities.LastItem as InlineContentControl;
        text = new WTextRange(document);
        text.ApplyCharacterFormat(dropDownCC.BreakCharacterFormat);
        text.Text = dropDownCC.ContentControlProperties.ContentControlListItems[2].DisplayText;
        dropDownCC.ParagraphItems.Add(text);
        //Fill the VB experience level.
        cellPara1 = row1.Cells[0].ChildEntities[9] as WParagraph;
        dropDownCC = cellPara1.ChildEntities.LastItem as InlineContentControl;
        text = new WTextRange(document);
        text.ApplyCharacterFormat(dropDownCC.BreakCharacterFormat);
        text.Text = dropDownCC.ContentControlProperties.ContentControlListItems[1].DisplayText;
        dropDownCC.ParagraphItems.Add(text);
        #endregion
        //Creates file stream.
        using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite))
        {
            //Saves the Word document to file stream.
            document.Save(outputFileStream, FormatType.Docx);
        }
    }
}
```

More information about the content controls can be found in this [documentation](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-content-controls) section.