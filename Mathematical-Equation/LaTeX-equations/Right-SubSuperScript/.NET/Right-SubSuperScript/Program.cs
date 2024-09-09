using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;


// Create a new Word document.
using WordDocument document = new WordDocument();

//Add one section and one paragraph to the document.
document.EnsureMinimal();

//Append an right subsuperscript equation using LaTeX.
document.LastParagraph.AppendMath(@"{100}_{40}^{20}");

//Save the Word document.
using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
document.Save(outputStream, FormatType.Docx);