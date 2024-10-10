using Syncfusion.DocIO.DLS;

using (WordDocument document = new WordDocument())
{
    //Initialize the paragraph where the bookmark will be inserted.
    IWParagraph paragraph = null;
    //Get the full path of the input HTML file.
    string htmlPage = Path.GetFullPath(@"Data/Input.html");
    //Read the HTML content from the input file.
    string htmlContent = File.ReadAllText(htmlPage);
    //Add a new section to the Word document.
    IWSection section = document.AddSection();
    //Add a paragraph to the section.
    paragraph = section.AddParagraph();
    //Insert a bookmark start at the current location in the paragraph.
    paragraph.AppendBookmarkStart("Index");
    //Insert a bookmark end at the current location in the paragraph.
    paragraph.AppendBookmarkEnd("Index");
    //Convert the HTML content into a WordDocumentPart.
    WordDocumentPart htmlDocumentPart = ConvertHTMLToWordDocumentPart(htmlContent);
    //Create an instance of BookmarksNavigator to navigate to the bookmark.
    BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(document);
    //Move the virtual cursor to the location of the bookmark named "Index".
    bookmarkNavigator.MoveToBookmark("Index");
    //Replace the bookmark content with the converted HTML content.
    bookmarkNavigator.ReplaceContent(htmlDocumentPart);
    //Save the modified document to a specified file path in DOCX format.
    using (FileStream outputStream = new FileStream(Path.GetFullPath("Output/Output.docx"), FileMode.Create, FileAccess.Write))
    {
        document.Save(outputStream, Syncfusion.DocIO.FormatType.Docx);
    }
}

/// <summary>
/// Converts an HTML string into a WordDocumentPart to be inserted into the Word document.
/// </summary>
static WordDocumentPart ConvertHTMLToWordDocumentPart(string html)
{
    //Create a new Word document.
    WordDocument tempDocument = new WordDocument();
    //Add minimal content to the document (ensures the document structure is valid).
    tempDocument.EnsureMinimal();
    //Append the HTML string to the last paragraph of the temporary Word document.
    tempDocument.LastParagraph.AppendHTML(html);
    //Create a new WordDocumentPart instance.
    WordDocumentPart wordDocumentPart = new WordDocumentPart();
    //Load the temporary document into the WordDocumentPart.
    wordDocumentPart.Load(tempDocument);
    //Return the WordDocumentPart containing the converted HTML content.
    return wordDocumentPart;
}
