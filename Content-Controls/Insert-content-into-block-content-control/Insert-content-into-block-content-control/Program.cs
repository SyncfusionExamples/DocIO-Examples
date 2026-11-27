using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

// Creates a new Word document.
using (WordDocument document = new WordDocument())
{
    // Adds new section to the document.
    IWSection section = document.AddSection();
    WTextBody textBody = section.Body;
    // Adds block content control into Word document.
    BlockContentControl blockContentControl = textBody.AddBlockContentControl(ContentControlType.RichText) as BlockContentControl;
    using (FileStream docStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read))
    {
        // Open the template Word document.
        using (WordDocument docxDocument = new WordDocument(docStream, FormatType.Docx))
        {
            // Insert the Word document into block content control.
            InsertContentIntoBlockContentControl(blockContentControl, docxDocument);
            using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
            {
                // Saves the Word document.
                document.Save(outputFileStream, FormatType.Docx);
            }
        }
    }
}

/// <summary>
/// Inserts the content of a specified Word document into a block content control.
/// </summary>
void InsertContentIntoBlockContentControl(BlockContentControl blockContentControl, WordDocument document)
{
    // Get the textbody of the another Word document and add it to the block content control.
    WTextBody textBody = document.LastSection.Body;
    foreach (IEntity entity in textBody.ChildEntities)
    {
        blockContentControl.TextBody.ChildEntities.Add(entity.Clone());
    }
}