//Open an existing document.
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

//Open an existing Word document from the specified file path.
FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open, FileAccess.Read);
//Create a new instance of the WordDocument class and load the document from the FileStream.
WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx);
//Retrieve the first section of the Word document.
WSection firstSection = document.Sections[0];
//Iterate through each section in the Word document starting from the second section.
for (int index = 1; index < document.Sections.Count; index++)
{
    //Copy the headers and footers from the first section to the current section.
    UpdateHeaderFooter(firstSection, document.Sections[index]);
}
//Remove the first section from the document.
document.Sections.RemoveAt(0);
//Create a FileStream to save the modified document to a new file at the specified path.
FileStream stream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.Write);
//Save the Word document to the stream in DOCX format.
document.Save(stream, FormatType.Docx);
//Close the WordDocument instance to release the resources.
document.Close();


///<summary>
///Copy the source section header and footer to destination section
/// </summary>
void UpdateHeaderFooter(WSection sourceSection, WSection destinationSection)
{
    //clear the destination section header and footer
    ClearHeaderFooter(destinationSection);

    //Add Headers
    for (int j = 0; j < sourceSection.HeadersFooters.Header.ChildEntities.Count; j++)
    {
        destinationSection.HeadersFooters.Header.ChildEntities.Add(sourceSection.HeadersFooters.Header.ChildEntities[j].Clone());
    }

    //Add Footers
    for (int j = 0; j < sourceSection.HeadersFooters.Footer.ChildEntities.Count; j++)
    {
        destinationSection.HeadersFooters.Footer.ChildEntities.Add(sourceSection.HeadersFooters.Footer.ChildEntities[j].Clone());
    }
}
///<summary>
///Clear all header and footer for the section
/// </summary>
void ClearHeaderFooter(WSection section)
{
    //Remove the first page header.
    section.HeadersFooters.FirstPageHeader.ChildEntities.Clear();
    //Remove the first page footer.
    section.HeadersFooters.FirstPageFooter.ChildEntities.Clear();
    //Remove the odd footer.
    section.HeadersFooters.OddFooter.ChildEntities.Clear();
    //Remove the odd header.
    section.HeadersFooters.OddHeader.ChildEntities.Clear();
    //Remove the even header.
    section.HeadersFooters.EvenHeader.ChildEntities.Clear();
    //Remove the even footer.
    section.HeadersFooters.EvenFooter.ChildEntities.Clear();
}