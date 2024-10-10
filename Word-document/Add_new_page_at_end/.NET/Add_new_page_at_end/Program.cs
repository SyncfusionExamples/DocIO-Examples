using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

// Open a file stream to read the existing Word document.
using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Adventure.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    // Initialize a WordDocument object with the opened file stream.
    using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
    {
        // Append a page break to the last paragraph of the document.
        document.LastParagraph.AppendBreak(BreakType.PageBreak);
        // Create a file stream for the output document to save the modified content.
        using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite))
        {
            // Save the modified Word document to the output file stream in Docx format.
            document.Save(outputFileStream, FormatType.Docx);
        }
    }
}
