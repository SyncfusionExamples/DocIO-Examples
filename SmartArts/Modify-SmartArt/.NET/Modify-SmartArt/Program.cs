using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using Syncfusion.Office;


// Open the input Word document as a file stream.
using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/EditSmartArtInput.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    // Load the Word document from the file stream.
    using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
    {
        //Gets the last paragraph in the document.
        WParagraph paragraph = document.LastParagraph;
        //Retrieves the SmartArt object from the paragraph.
        WSmartArt smartArt = paragraph.ChildEntities[0] as WSmartArt;
        //Sets the background fill type of the SmartArt to solid.
        smartArt.Background.FillType = OfficeShapeFillType.Solid;
        //Sets the background color of the SmartArt.
        smartArt.Background.SolidFill.Color = Color.FromArgb(255, 242, 169, 132);
        //Gets the first node of the SmartArt.
        IOfficeSmartArtNode node = smartArt.Nodes[0];
        //Modifies the text content of the first node.
        node.TextBody.Text = "Goals";
        //Retrieves the first shape of the node.
        IOfficeSmartArtShape shape = node.Shapes[0];
        //Sets the fill color of the shape.
        shape.Fill.SolidFill.Color = Color.FromArgb(255, 160, 43, 147);
        //Sets the line format color of the shape.
        shape.LineFormat.Fill.SolidFill.Color = Color.FromArgb(255, 160, 43, 147);
        //Gets the first child node of the current node.
        IOfficeSmartArtNode childNode = node.ChildNodes[0];
        //Modifies the text content of the child node.
        childNode.TextBody.Text = "Set clear goals to the team.";
        //Sets the line format color of the first shape in the child node.
        childNode.Shapes[0].LineFormat.Fill.SolidFill.Color = Color.FromArgb(255, 160, 43, 147);

        //Retrieves the second node in the SmartArt and updates its text content.
        node = smartArt.Nodes[1];
        node.TextBody.Text = "Progress";

        //Retrieves the third node in the SmartArt and updates its text content.
        node = smartArt.Nodes[2];
        node.TextBody.Text = "Result";
        //Retrieves the first shape of the third node.
        shape = node.Shapes[0];
        //Sets the fill color of the shape.
        shape.Fill.SolidFill.Color = Color.FromArgb(255, 78, 167, 46);
        //Sets the line format color of the shape.
        shape.LineFormat.Fill.SolidFill.Color = Color.FromArgb(255, 78, 167, 46);
        //Sets the line format color of the first shape in the child node.
        node.ChildNodes[0].Shapes[0].LineFormat.Fill.SolidFill.Color = Color.FromArgb(255, 78, 167, 46);

        // Create a file stream for saving the modified document.
        using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
        {
            // Save the modified Word document to the output file stream.
            document.Save(outputFileStream, FormatType.Docx);
        }
    }
}
