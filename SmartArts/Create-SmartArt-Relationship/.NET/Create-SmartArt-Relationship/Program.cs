using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;


WordDocument document = new WordDocument();
IWSection section = document.AddSection();
// Retrieves the first paragraph and add text.
IWParagraph paragraph = section.AddParagraph();
paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
IWTextRange textRange1 = paragraph.AppendText("Remote Work vs Office Work");
textRange1.CharacterFormat.FontSize = 28f;

paragraph = section.AddParagraph();
paragraph = section.AddParagraph();
paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;

// Add SmartArt with "Counter Balance Arrows" layout
WSmartArt counterBalanceArrowsSmartArt = paragraph.AppendSmartArt(OfficeSmartArtType.CounterBalanceArrows, 432, 252);
// Add the "Remote Work" phase node
IOfficeSmartArtNode remoteWorkNode = counterBalanceArrowsSmartArt.Nodes[0];
remoteWorkNode.TextBody.Text = "Remote Work";
remoteWorkNode.TextBody.Paragraphs[0].TextParts[0].Font.FontSize = 19f;
remoteWorkNode.ChildNodes.Add();
remoteWorkNode.ChildNodes.Add();
AddSmartArtChildNode(remoteWorkNode, "Flexibility", "Work-Life Balance", 15f);

// Add the "Office Work" phase node
IOfficeSmartArtNode officeWorkNode = counterBalanceArrowsSmartArt.Nodes[1];
officeWorkNode.TextBody.Text = "Office Work";
officeWorkNode.TextBody.Paragraphs[0].TextParts[0].Font.FontSize = 19f;
officeWorkNode.ChildNodes.Add();
officeWorkNode.ChildNodes.Add();
AddSmartArtChildNode(officeWorkNode, "Collaboration", "Structured Environment", 15f);

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
