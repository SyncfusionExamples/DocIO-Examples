using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using Syncfusion.Office;

// Create a new Word document instance.
WordDocument document = new WordDocument();
// Add a new section to the document.
IWSection section = document.AddSection();
// Add a paragraph to the section.
IWParagraph paragraph = section.AddParagraph();
// Append a SmartArt object of type "Alternating Hexagons" to the paragraph with specified dimensions.
WSmartArt smartArt = paragraph.AppendSmartArt(OfficeSmartArtType.AlternatingHexagons, 432, 252);
// Add a new node to the SmartArt.
IOfficeSmartArtNode newNode = smartArt.Nodes.Add();
// Set text content for the newly added SmartArt node.
newNode.TextBody.AddParagraph("New main node added.");
//Creates file stream.
using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
{
    //Saves the Word document to file stream.
    document.Save(outputFileStream, FormatType.Docx);
}