using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;


// Create a new Word document.
using (WordDocument document = new WordDocument())
{
    //Add one section and one paragraph to the document.
    document.EnsureMinimal();

    //Append an accent equation with bold using LaTeX.
    document.LastParagraph.AppendMath(@"\dot{\mathbf{a}}");
    //Append an accent equation with bold-italic using LaTeX.
    document.LastSection.AddParagraph().AppendMath(@"\dot{\mathbit{a}}");

    //Save a Word document.
    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
    {
        //Saves the Word document to file stream.
        document.Save(outputFileStream, FormatType.Docx);
    }
}
