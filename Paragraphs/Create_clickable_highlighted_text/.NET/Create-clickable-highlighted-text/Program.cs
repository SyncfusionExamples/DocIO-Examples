using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;

//Create a new Word document
WordDocument document = new WordDocument();

// Add a section
IWSection section = document.AddSection();

// Add introductory paragraph
IWParagraph para = section.AddParagraph();
para.AppendText("Student Handbook\n\n");
para.AppendText("For attendance rules, see ");

// Add bookmark hyperlink
IWField hyperlink = para.AppendHyperlink("Attendance","Attendance Policy",HyperlinkType.Bookmark);

// Highlight the hyperlink text
WTextRange textRange =hyperlink.OwnerParagraph.ChildEntities[hyperlink.OwnerParagraph.ChildEntities.Count - 2] as WTextRange;

if (textRange != null)
{
    textRange.CharacterFormat.HighlightColor = Color.Yellow;
}

para.AppendText(".");

// Add some content before the destination
section.AddParagraph().AppendText("Student Code of Conduct");
section.AddParagraph().AppendText("Examination Rules");
section.AddParagraph().AppendText("Library Guidelines");

// Add bookmark destination
IWParagraph destination = section.AddParagraph();
destination.AppendBookmarkStart("Attendance");
destination.AppendText("Attendance Policy");
destination.AppendBookmarkEnd("Attendance");

// Add attendance rules as bullet list
IWParagraph bullet1 = section.AddParagraph();
bullet1.ListFormat.ApplyDefBulletStyle();
bullet1.AppendText("Students must maintain at least 75% attendance.");

IWParagraph bullet2 = section.AddParagraph();
bullet2.ListFormat.ApplyDefBulletStyle();
bullet2.AppendText("Students should attend all mandatory classes.");

IWParagraph bullet3 = section.AddParagraph();
bullet3.ListFormat.ApplyDefBulletStyle();
bullet3.AppendText("Medical leave must be supported by valid documents.");

// Save document
document.Save(@"../../../Output/Output.docx", FormatType.Docx);

// Close document
document.Close();