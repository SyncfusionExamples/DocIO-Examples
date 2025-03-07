using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;


WordDocument document = new WordDocument();
IWSection section = document.AddSection();
// Retrieves the first paragraph and add text.
IWParagraph paragraph = section.AddParagraph();
paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
IWTextRange textRange1 = paragraph.AppendText("Company Organizational Structure");
textRange1.CharacterFormat.FontSize = 28f;

paragraph = section.AddParagraph();
paragraph = section.AddParagraph();
paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;

// Add SmartArt with "Hierarchy" layout
WSmartArt hierarchySmartArt = paragraph.AppendSmartArt(OfficeSmartArtType.Hierarchy, 432, 252);

// Configure the "Manager" node and its hierarchy
IOfficeSmartArtNode manager = hierarchySmartArt.Nodes[0];
manager.TextBody.Text = "Manager";
manager.TextBody.Paragraphs[0].TextParts[0].Font.FontSize = 20f;
manager.ChildNodes[0].TextBody.Text = "Team Lead 1";
manager.ChildNodes[0].TextBody.Paragraphs[0].TextParts[0].Font.FontSize = 20f;
manager.ChildNodes[0].ChildNodes[0].TextBody.Text = "Employee 1";
manager.ChildNodes[0].ChildNodes[0].TextBody.Paragraphs[0].TextParts[0].Font.FontSize = 20f;
manager.ChildNodes[0].ChildNodes[1].TextBody.Text = "Employee 2";
manager.ChildNodes[0].ChildNodes[1].TextBody.Paragraphs[0].TextParts[0].Font.FontSize = 20f;
manager.ChildNodes[1].TextBody.Text = "Team Lead 2";
manager.ChildNodes[1].TextBody.Paragraphs[0].TextParts[0].Font.FontSize = 20f;
manager.ChildNodes[1].ChildNodes[0].TextBody.Text = "Employee 3";
manager.ChildNodes[1].ChildNodes[0].TextBody.Paragraphs[0].TextParts[0].Font.FontSize = 20f;

//Creates file stream.
using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
{
    //Saves the Word document to file stream.
    document.Save(outputFileStream, FormatType.Docx);
}
