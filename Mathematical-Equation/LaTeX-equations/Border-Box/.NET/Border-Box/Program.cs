using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;


// Create a new Word document.
using WordDocument document = new WordDocument();

//Add one section and one paragraph to the document.
document.EnsureMinimal();

//Append an border box equation using LaTeX.
document.LastParagraph.AppendMath(@"\boxed{x^2 + y^2 = z^2}");

//Save the Word document.
using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
document.Save(outputStream, FormatType.Docx);