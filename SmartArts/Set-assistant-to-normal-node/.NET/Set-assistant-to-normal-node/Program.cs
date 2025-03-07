using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using Syncfusion.Office;
using System.IO;

//Create a new Word document instance.
WordDocument document = new WordDocument();
//Add a new section to the document.
IWSection section = document.AddSection();
//Add a paragraph to the section.
IWParagraph paragraph = section.AddParagraph();
//Add a SmartArt to the paragraph at the specified size.
WSmartArt smartArt = paragraph.AppendSmartArt(OfficeSmartArtType.OrganizationChart, 432, 252);
//Traverse through all nodes of the SmartArt.
foreach (IOfficeSmartArtNode node in smartArt.Nodes)
{
    //Check if the node is assistant or not.
    if (node.IsAssistant)
        //Set the assistant node to false.
        node.IsAssistant = false;
}
// Create a file stream to save the modified document.
using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
{
    // Save the modified Word document to the output file.
    document.Save(outputFileStream, FormatType.Docx);
}

