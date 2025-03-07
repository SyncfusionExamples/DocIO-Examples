using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;

// Open the input Word document as a file stream.
using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    // Load the Word document from the file stream.
    using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
    {
        // Get the last paragraph in the document.
        WParagraph paragraph = document.LastParagraph;

        // Iterate through all child elements in the paragraph.
        for (int i = 0; i < paragraph.ChildEntities.Count; i++)
        {
            // Check if the current child entity is a SmartArt object.
            if (paragraph.ChildEntities[i] is WSmartArt)
            {
                // Traverse through all nodes inside the SmartArt.
                foreach (IOfficeSmartArtNode node in (paragraph.ChildEntities[i] as WSmartArt).Nodes)
                {
                    // Check if the node contains the text "Inquiry".
                    if (node.TextBody.Text == "Inquiry")
                    {
                        // Update the text content of the node to "New Content".
                        node.TextBody.Paragraphs[0].TextParts[0].Text = "New Content";
                    }
                }
            }
        }

        //Creates file stream.
        using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
        {
            //Saves the Word document to file stream.
            document.Save(outputFileStream, FormatType.Docx);
        }
    }
}
