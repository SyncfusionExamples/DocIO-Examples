using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using Syncfusion.Office;
using System.IO;

//Create a new Word document instance.
WordDocument wordDocument = new WordDocument();
//Add one section and one paragraph to the document.
wordDocument.EnsureMinimal();
//Add a SmartArt to the paragraph at the specified size.
WSmartArt smartArt = wordDocument.LastParagraph.AppendSmartArt(OfficeSmartArtType.OrganizationChart, 640, 426);
//Traverse through all nodes of the SmartArt.
foreach (IOfficeSmartArtNode node in smartArt.Nodes)
{
    foreach (IOfficeSmartArtNode childNode in node.ChildNodes)
    {
        //Check if the node is assistant or not.
        if (childNode.IsAssistant)
            //Set the assistant node to false.
            childNode.IsAssistant = false;
    }
}
// Create a file stream to save the modified document.
using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
{
    // Save the modified Word document to the output file.
    document.Save(outputFileStream, FormatType.Docx);
}

