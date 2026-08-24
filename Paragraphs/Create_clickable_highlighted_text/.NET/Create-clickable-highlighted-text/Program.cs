using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;

WordDocument document = new WordDocument();
IWSection section = document.AddSection();

// Paragraph containing highlighted hyperlink word
IWParagraph para = section.AddParagraph();

para.AppendText("For detailed information, refer to ");

IWField hyperlink = para.AppendHyperlink(
    "Section5",                  // Bookmark name
    "Section 5",                 // Display text
    HyperlinkType.Bookmark);


WTextRange textRange = hyperlink.OwnerParagraph.ChildEntities[
    hyperlink.OwnerParagraph.ChildEntities.Count - 2] as WTextRange;

if (textRange != null)
{
    textRange.CharacterFormat.HighlightColor = Color.Yellow;
}

para.AppendText(" in this document.");

//Add some content before destination
for (int i = 0; i < 5; i++)
{
    para = section.AddParagraph();
    para.AppendBookmarkStart($"Section{i}");
    para.AppendText($"Sample paragraph {i + 1}");
    para.AppendBookmarkEnd($"Section{i}");
}

// Bookmark destination
IWParagraph destinationPara = section.AddParagraph();

//Bookmark start
destinationPara.AppendBookmarkStart("Section5");

//Destination content
destinationPara.AppendText("Section 5 - Detailed Information");

//Bookmark end
destinationPara.AppendBookmarkEnd("Section5");
section.AddParagraph().AppendText("End");
//Save document
document.Save(@"../../../Output/Output.docx", FormatType.Docx);
document.Close();