using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

using (FileStream inputFileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read))
{
    // Opens the template Word docuemnt.
    using (WordDocument document = new WordDocument(inputFileStream, FormatType.Docx))
    {
        // Deletes content from the 2nd to the 6th index in the text body of the Word document.
        DeleteContentInWordDocument(2, 6, document);
        using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
        {
            // Saves the modified Word document to the output file stream.
            document.Save(outputFileStream, FormatType.Docx);
        }
    }
}

/// <summary>
/// Deletes content from the specified start index to the end index within the text body of the Word document.
/// </summary>
 void DeleteContentInWordDocument(int startIndex, int endIndex, WordDocument document)
{
    // Retrieves the text body of the last section in the Word document.
    WTextBody body = document.LastSection.Body;
    // Ensures the start and end indices are within valid bounds.
    if (startIndex >= 0 && endIndex < body.ChildEntities.Count && startIndex <= endIndex)
    {
        // Removes items from endIndex to startIndex in reverse order to prevent index shifting issues.
        for (int index = endIndex; index >= startIndex; index--)
        {
            body.ChildEntities.RemoveAt(index);
        }
    }
}
