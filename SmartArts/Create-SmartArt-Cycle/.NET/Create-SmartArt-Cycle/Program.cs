using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;


WordDocument document = new WordDocument();
IWSection section = document.AddSection();
// Retrieves the first paragraph and add text.
IWParagraph paragraph = section.AddParagraph();
paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
IWTextRange textRange1 = paragraph.AppendText("Customer Service Cycle");
textRange1.CharacterFormat.FontSize = 28f;
paragraph = section.AddParagraph();
paragraph = section.AddParagraph();
paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;

// Add SmartArt with "Block Cycle" layout
WSmartArt blockCycleSmartArt = paragraph.AppendSmartArt(OfficeSmartArtType.BlockCycle, 432, 252);

// Add the "Inquiry" phase node
IOfficeSmartArtNode inquiryNode = blockCycleSmartArt.Nodes[0];
inquiryNode.TextBody.Text = "Inquiry";
inquiryNode.TextBody.Paragraphs[0].TextParts[0].Font.FontSize = 15f;

// Add the "Response" phase node
IOfficeSmartArtNode responseNode = blockCycleSmartArt.Nodes[1];
responseNode.TextBody.Text = "Response";
responseNode.TextBody.Paragraphs[0].TextParts[0].Font.FontSize = 15f;

// Add the "Resolution" phase node
IOfficeSmartArtNode resolutionNode = blockCycleSmartArt.Nodes[2];
resolutionNode.TextBody.Text = "Resolution";
resolutionNode.TextBody.Paragraphs[0].TextParts[0].Font.FontSize = 15f;

// Add the "Feedback" phase node
IOfficeSmartArtNode feedBackNode = blockCycleSmartArt.Nodes[3];
feedBackNode.TextBody.Text = "Feedback";
feedBackNode.TextBody.Paragraphs[0].TextParts[0].Font.FontSize = 15f;

// Add the "Follow-up" phase node
IOfficeSmartArtNode followupNode = blockCycleSmartArt.Nodes[4];
followupNode.TextBody.Text = "Follow-up";
followupNode.TextBody.Paragraphs[0].TextParts[0].Font.FontSize = 15f;

//Creates file stream.
using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
{
    //Saves the Word document to file stream.
    document.Save(outputFileStream, FormatType.Docx);
}
