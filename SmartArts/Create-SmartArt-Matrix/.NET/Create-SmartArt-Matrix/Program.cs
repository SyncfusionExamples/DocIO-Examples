using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;


WordDocument document = new WordDocument();
IWSection section = document.AddSection();
// Retrieves the first paragraph and add text.
IWParagraph paragraph = section.AddParagraph();
paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
IWTextRange textRange1 = paragraph.AppendText("Marketing Campaign Process");
textRange1.CharacterFormat.FontSize = 28f;

paragraph = section.AddParagraph();
paragraph = section.AddParagraph();
paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;

// Add SmartArt with "Grid Matrix" layout
WSmartArt gridMatrixSmartArt = paragraph.AppendSmartArt(OfficeSmartArtType.GridMatrix, 432, 252);

// Add the "Planning" phase node
IOfficeSmartArtNode planningNode = gridMatrixSmartArt.Nodes[0];
planningNode.TextBody.Text = "Planning";
planningNode.TextBody.Paragraphs[0].TextParts[0].Font.FontSize = 13f;
planningNode.ChildNodes.Add();
planningNode.ChildNodes.Add();
AddSmartArtChildNode(planningNode, "Define goals and target audience.", "Identify key messaging and channels.", 10f);

// Add the "Execution" phase node
IOfficeSmartArtNode executionNode = gridMatrixSmartArt.Nodes[1];
executionNode.TextBody.Text = "Execution";
executionNode.TextBody.Paragraphs[0].TextParts[0].Font.FontSize = 13f;
executionNode.ChildNodes.Add();
executionNode.ChildNodes.Add();
AddSmartArtChildNode(executionNode, "Create content and implement strategies.", "Launch campaigns across chosen platforms.", 10f);

// Add the "Monitoring" phase node
IOfficeSmartArtNode monitoringNode = gridMatrixSmartArt.Nodes[2];
monitoringNode.TextBody.Text = "Monitoring";
monitoringNode.TextBody.Paragraphs[0].TextParts[0].Font.FontSize = 13f;
monitoringNode.ChildNodes.Add();
monitoringNode.ChildNodes.Add();
AddSmartArtChildNode(monitoringNode, "Track performance and engagement.", "Collect data and identify trends.", 10f);

// Add the "Optimization" phase node
IOfficeSmartArtNode optimizingNode = gridMatrixSmartArt.Nodes[3];
optimizingNode.TextBody.Text = "Optimization";
optimizingNode.TextBody.Paragraphs[0].TextParts[0].Font.FontSize = 13f;
optimizingNode.ChildNodes.Add();
optimizingNode.ChildNodes.Add();
AddSmartArtChildNode(optimizingNode, "Adjust strategies based on insights.", "Fine-tune campaigns for better results.", 10f);

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
