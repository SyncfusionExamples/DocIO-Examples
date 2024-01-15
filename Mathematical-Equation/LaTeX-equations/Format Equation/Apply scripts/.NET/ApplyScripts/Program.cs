using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;


// Create a new Word document.
using WordDocument document = new WordDocument();

//Add one section and one paragraph to the document.
document.EnsureMinimal();

//Append an accent equation with DoubleStruck font using LaTeX.
document.LastParagraph.AppendMath(@"\dot{\mathbb{a}}");
//Append an accent equation with Fraktur font using LaTeX.
document.LastSection.AddParagraph().AppendMath(@"\dot{\mathfrak{a}}");
//Append an accent equation with SansSerif font using LaTeX.
document.LastSection.AddParagraph().AppendMath(@"\dot{\mathsf{a}}");
//Append an accent equation with Script using LaTeX.
document.LastSection.AddParagraph().AppendMath(@"\dot{\mathcal{a}}");
//Append a box equation with Script using LaTeX.
document.LastSection.AddParagraph().AppendMath(@"\dot{\mathscr{a}}");

//Save the Word document.
using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
document.Save(outputStream, FormatType.Docx);