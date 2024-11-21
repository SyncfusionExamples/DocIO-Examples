# Format table in a Word document using C#

The Syncfusion [.NET Word Library](https://www.syncfusion.com/document-processing/word-framework/net/word-library) (DocIO) enables you to create, read, and edit Word documents programmatically without Microsoft Word or interop dependencies. Using this library, you can **format tables in a Word document** using C#.

## Steps to format a table programmatically

Step 1: Create a new .NET Core console application project.

Step 2: Install the [Syncfusion.DocIO.Net.Core](https://www.nuget.org/packages/Syncfusion.DocIO.Net.Core) NuGet package as a reference to your project from [NuGet.org](https://www.nuget.org/).

Step 3: Include the following namespaces in the Program.cs file.

```csharp
using Syncfusion.DocIO; 
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.IO;
```

Step 4: Add the following code snippet in Program.cs file to format a table in the Word document.

```csharp
using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    //Creates an instance of WordDocument class (Empty Word Document).
    using (WordDocument document = new WordDocument())
    {
        //Opens an existing Word document into DocIO instance.
        document.Open(fileStreamPath, FormatType.Docx);
        //Accesses the instance of the first section in the Word document.
        WSection section = document.Sections[0];
        //Accesses the instance of the first table in the section.
        WTable table = section.Tables[0] as WTable;
        //Specifies the title for the table.
        table.Title = "PriceDetails";
        //Specifies the description of the table.
        table.Description = "This table shows the price details of various fruits";
        //Specifies the left indent of the table.
        table.IndentFromLeft = 50;
        //Specifies the background color of the table.
        table.TableFormat.BackColor = Color.FromArgb(192, 192, 192);
        //Specifies the horizontal alignment of the table.
        table.TableFormat.HorizontalAlignment = RowAlignment.Left;
        //Specifies the left, right, top and bottom padding of all the cells in the table.
        table.TableFormat.Paddings.All = 10;
        //Specifies the auto resize of table to automatically resize all cell width based on its content.
        table.TableFormat.IsAutoResized = true;
        //Specifies the table top, bottom, left and right border line width.
        table.TableFormat.Borders.LineWidth = 2f;
        //Specifies the table horizontal border line width.
        table.TableFormat.Borders.Horizontal.LineWidth = 2f;
        //Specifies the table vertical border line width.
        table.TableFormat.Borders.Vertical.LineWidth = 2f;
        //Specifies the tables top, bottom, left and right border color.
        table.TableFormat.Borders.Color = Color.Red;
        //Specifies the table Horizontal border color.
        table.TableFormat.Borders.Horizontal.Color = Color.Red;
        //Specifies the table vertical border color.
        table.TableFormat.Borders.Vertical.Color = Color.Red;
        //Specifies the table borders border type.
        table.TableFormat.Borders.BorderType = BorderStyle.Double;
        //Accesses the instance of the first row in the table.
        WTableRow row = table.Rows[0];
        //Specifies the row height.
        row.Height = 20;
        //Specifies the row height type.
        row.HeightType = TableRowHeightType.AtLeast;
        //Creates file stream.
        using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite))
        {
            //Saves the Word document to file stream.
            document.Save(outputFileStream, FormatType.Docx);
        }
    }
}
```

More information about the table formatting can be found in this [documentation](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-tables#apply-formatting-to-table-row-and-cell) section.