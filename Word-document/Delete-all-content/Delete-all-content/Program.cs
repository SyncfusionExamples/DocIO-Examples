using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

using (FileStream inputFileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read))
{
    // Opens the template Word docuemnt.
    using (WordDocument document = new WordDocument(inputFileStream, FormatType.Docx))
    {
        //Delete all the content in the Word document 
        DeleteAllContentInWordDocument(document);
        using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
        {
            // Saves the modified Word document to the output file stream.
            document.Save(outputFileStream, FormatType.Docx);
        }
    }
}

/// <summary>
/// Deletes all content from every section of the Word document.
/// </summary>
void DeleteAllContentInWordDocument(WordDocument document)
{
    // Iterate through all sections in the Word document.
    foreach (WSection section in document.Sections)
    {
        // Access the body of the current section.
        WTextBody body = section.Body;
        // Remove all child entities from the body.
        body.ChildEntities.Clear();
    }
}
