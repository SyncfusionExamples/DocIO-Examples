using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;


// Create a new Word document.
using (WordDocument document = new WordDocument())
{
    //Add one section and one paragraph to the document.
    document.EnsureMinimal();

    //Append an accent equation as normal text using LaTeX.
    document.LastParagraph.AppendMath(@"\dot{\mathrm{a}}");

    //Save a Word document to the MemoryStream.
    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
    {
        //Saves the Word document to file stream.
        document.Save(outputFileStream, FormatType.Docx);
    }
}
