using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;


WordDocument document = new WordDocument();
IWSection section = document.AddSection();
// Retrieves the first paragraph and add text.
IWParagraph paragraph = section.AddParagraph();
paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
IWTextRange textRange = paragraph.AppendText("Project Planning List");
textRange.CharacterFormat.FontSize = 28f;
paragraph = section.AddParagraph();
paragraph = section.AddParagraph();
paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;

// Add SmartArt with "Vertical Chevron List" layout.
WSmartArt verticalChevronListSmartArt = paragraph.AppendSmartArt(OfficeSmartArtType.VerticalChevronList, 432, 252);

// Add the "Planning" phase node.
IOfficeSmartArtNode planningNode = verticalChevronListSmartArt.Nodes[0];
planningNode.TextBody.Text = "Planning";
planningNode.TextBody.Paragraphs[0].TextParts[0].Font.FontSize = 15f;
AddSmartArtChildNode(planningNode, "Set clear objectives.", "Allocate resources effectively.", 23f);

// Add the "Execution" phase node.
IOfficeSmartArtNode executionNode = verticalChevronListSmartArt.Nodes[1];
executionNode.TextBody.Text = "Execution";
executionNode.TextBody.Paragraphs[0].TextParts[0].Font.FontSize = 15f;
AddSmartArtChildNode(executionNode, "Assign tasks to the team.", "Track progress regularly.", 23f);

// Add the "Review" phase node.
IOfficeSmartArtNode reviewNode = verticalChevronListSmartArt.Nodes[2];
reviewNode.TextBody.Text = "Review";
reviewNode.TextBody.Paragraphs[0].TextParts[0].Font.FontSize = 15f;
AddSmartArtChildNode(reviewNode, "Analyze outcomes.", "Identify lessons learned.", 23f);

//Creates file stream.
using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
{
    //Saves the Word document to file stream.
    document.Save(outputFileStream, FormatType.Docx);
}


/// <summary>
/// Adds two child nodes to a given SmartArt node and applies formatting.
/// </summary>
void AddSmartArtChildNode(IOfficeSmartArtNode node, string childText1, string childText2, float fontSize)
{
    node.ChildNodes[0].TextBody.Text = childText1;
    node.ChildNodes[1].TextBody.Text = childText2;
    for (int i = 0; i < node.ChildNodes.Count; i++)
    {
        node.ChildNodes[i].TextBody.Paragraphs[0].TextParts[0].Font.FontSize = fontSize;
    }
}
