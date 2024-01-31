using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;

// Create a new Word document.
using (WordDocument document = new WordDocument())
{
    //Add one section and one paragraph to the document.
    document.EnsureMinimal();
    //Append an border box equation using LaTeX.
    WMath math = document.LastParagraph.AppendMath(@"\boxed{{x}^{2}+{y}^{2}={z}^{2}}");
    //Apply math justification.
    math.MathParagraph.Justification = MathJustification.Left;
    using (FileStream outputFileStream = new FileStream("Output.docx", FileMode.Create, FileAccess.ReadWrite))
    {
        //Save Word document.
        document.Save(outputFileStream, FormatType.Docx);
    }
}
