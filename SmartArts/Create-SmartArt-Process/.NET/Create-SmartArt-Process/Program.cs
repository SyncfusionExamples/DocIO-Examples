using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;


WordDocument document = new WordDocument();
IWSection section = document.AddSection();
// Retrieves the first paragraph and add text.
IWParagraph paragraph = section.AddParagraph();
paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
IWTextRange textRange1 = paragraph.AppendText("Healthy Lifestyle Journey");
textRange1.CharacterFormat.FontSize = 28f;
textRange1.CharacterFormat.FontName = "Times New Roman";
paragraph = section.AddParagraph();
paragraph = section.AddParagraph();
paragraph.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;

// Add SmartArt with "Segmented Process" layout
WSmartArt segmentedProcessSmartArt = paragraph.AppendSmartArt(OfficeSmartArtType.SegmentedProcess, 432, 252);

// Add the "Start" phase node
IOfficeSmartArtNode startProcess = segmentedProcessSmartArt.Nodes[0];
startProcess.TextBody.Text = "Start";
startProcess.TextBody.Paragraphs[0].TextParts[0].Font.FontSize = 15f;
AddSmartArtChildNode(startProcess, "Begin exercise and eat balanced meals.", "Stay hydrated and prioritize sleep.", 12f);

// Add the "Track" phase node
IOfficeSmartArtNode trackProcess = segmentedProcessSmartArt.Nodes[1];
trackProcess.TextBody.Text = "Track";
trackProcess.TextBody.Paragraphs[0].TextParts[0].Font.FontSize = 15f;
AddSmartArtChildNode(trackProcess, "Record physical activity and food intake.", "Use a fitness app or journal.", 12f);

// Add the "Sustain" phase node
IOfficeSmartArtNode sustainProcess = segmentedProcessSmartArt.Nodes[2];
sustainProcess.TextBody.Text = "Sustain";
sustainProcess.TextBody.Paragraphs[0].TextParts[0].Font.FontSize = 15f;
AddSmartArtChildNode(sustainProcess, "Maintain healthy habits consistently.", "Set goals for continuous improvement.", 12f);

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
