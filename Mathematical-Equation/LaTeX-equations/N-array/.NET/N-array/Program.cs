using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;


// Create a new Word document.
using WordDocument document = new WordDocument();

//Add one section and one paragraph to the document.
document.EnsureMinimal();

//Append an N-array equation using LaTeX.
document.LastParagraph.AppendMath(@"\sum a");

//Save the Word document.
using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
document.Save(outputStream, FormatType.Docx);