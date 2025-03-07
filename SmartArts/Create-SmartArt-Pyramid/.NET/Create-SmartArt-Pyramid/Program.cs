using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;

WordDocument document = new WordDocument();
IWSection section = document.AddSection();
// Retrieves the first paragraph and add text.
IWParagraph paragraph = section.AddParagraph();
paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
IWTextRange textRange1 = paragraph.AppendText("Personal Growth");
textRange1.CharacterFormat.FontSize = 28f;

paragraph = section.AddParagraph();
paragraph = section.AddParagraph();
paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;

// Add SmartArt with "Basic Pyramid" layout
WSmartArt basicPyramidSmartArt = paragraph.AppendSmartArt(OfficeSmartArtType.BasicPyramid, 432, 252);

// Add the "Achievement" phase node
IOfficeSmartArtNode achievementNode = basicPyramidSmartArt.Nodes[0];
achievementNode.TextBody.Text = "Achievement";
achievementNode.TextBody.Paragraphs[0].TextParts[0].Font.FontSize = 26f;

// Add the "Skill Development" phase node
IOfficeSmartArtNode SkilldevelopmentNode = basicPyramidSmartArt.Nodes[1];
SkilldevelopmentNode.TextBody.Text = "Skill Development";
SkilldevelopmentNode.TextBody.Paragraphs[0].TextParts[0].Font.FontSize = 26f;

// Add the "Self-Awareness" phase node
IOfficeSmartArtNode selfAwarenessNode = basicPyramidSmartArt.Nodes[2];
selfAwarenessNode.TextBody.Text = "Self-Awareness";
selfAwarenessNode.TextBody.Paragraphs[0].TextParts[0].Font.FontSize = 26f;

//Creates file stream.
using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
{
    //Saves the Word document to file stream.
    document.Save(outputFileStream, FormatType.Docx);
}

