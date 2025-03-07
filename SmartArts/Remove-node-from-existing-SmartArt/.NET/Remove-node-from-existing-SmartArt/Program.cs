using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using Syncfusion.Office;
using System.IO;

// Open an existing Word document from the specified file path in read mode.
using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    // Load the Word document into memory.
    using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
    {
        //Retrieve the first SmartArt object from the last paragraph in the document.
        WSmartArt smartArt = document.LastParagraph.ChildEntities[0] as WSmartArt;
        //Remove a node at the specified index.
        smartArt.Nodes.RemoveAt(1);

        // Create a file stream to save the modified document.
        using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
        {
            // Save the modified Word document to the output file.
            document.Save(outputFileStream, FormatType.Docx);
        }
    }
}
