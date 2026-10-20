# Control bookmark end position

This sample shows how to control bookmark-end positions in a Word document using the [.NET Word Library](https://www.syncfusion.com/document-sdk/net-word-library) (Essential DocIO).

The sample demonstrates these `BookmarkEnd` APIs:

- `IsAfterParagraphMark`
- `IsAfterCellMark`
- `IsAfterRowMark`
- `IsAfterTableMark`

## Prerequisites

- [.NET SDK 8.0](https://dotnet.microsoft.com/download) or later
- [Syncfusion.DocIO.Net.Core](https://www.nuget.org/packages/Syncfusion.DocIO.Net.Core) NuGet package

## How to run the sample

1. Open `Control-bookmark-end-position.sln` in Visual Studio, or open a terminal in the `Control-bookmark-end-position` project folder.
2. Restore packages and run the project:

```
dotnet run
```

## Expected output

The sample creates `Output/BookmarkEndPositions.docx` with four bookmarks:

- `ParagraphBookmark` ends after the paragraph mark.
- `CellBookmark` ends after the table-cell mark.
- `RowBookmark` ends after the table-row mark.
- `TableBookmark` ends after the table mark.

More information about bookmarks can be found in this [documentation](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-bookmarks) section.
