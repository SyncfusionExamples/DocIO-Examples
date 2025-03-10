using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using Syncfusion.Office;
using System.IO;

// Open an existing Word document from the specified file path.
using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    // Load the Word document into memory.
    using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
    {
        // Retrieve the last paragraph in the document.
        WParagraph paragraph = document.LastParagraph;

        // Iterate through all child elements in the paragraph.
        for (int i = 0; i < paragraph.ChildEntities.Count; i++)
        {
            // Check if the current child entity is a SmartArt object.
            if (paragraph.ChildEntities[i] is WSmartArt)
            {
                // Remove the SmartArt object from the paragraph.
                paragraph.Items.RemoveAt(i);
                i--; 
            }
        }

        // Create a file stream to save the modified document.
        using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
        {
            // Save the modified Word document to the output file.
            document.Save(outputFileStream, FormatType.Docx);
        }
    }
}
