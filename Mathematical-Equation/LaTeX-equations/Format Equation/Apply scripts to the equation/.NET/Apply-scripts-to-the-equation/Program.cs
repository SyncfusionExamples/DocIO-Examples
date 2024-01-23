using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;


// Create a new Word document.
using (WordDocument document = new WordDocument())
{
    //Add one section and one paragraph to the document.
    document.EnsureMinimal();

    //Append an accent equation with Double-Struck font using LaTeX.
    document.LastParagraph.AppendMath(@"\dot{\mathbb{a}}");
    //Append an accent equation with Fraktur font using LaTeX.
    document.LastSection.AddParagraph().AppendMath(@"\dot{\mathfrak{a}}");
    //Append an accent equation with Sans Serif font using LaTeX.
    document.LastSection.AddParagraph().AppendMath(@"\dot{\mathsf{a}}");
    //Append an accent equation with Script using LaTeX.
    document.LastSection.AddParagraph().AppendMath(@"\dot{\mathcal{a}}");
    //Append an accent equation with Script using LaTeX.
    document.LastSection.AddParagraph().AppendMath(@"\dot{\mathscr{a}}");

    //Save a Word document
    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
    {
        //Saves the Word document to file stream.
        document.Save(outputFileStream, FormatType.Docx);
    }
}
