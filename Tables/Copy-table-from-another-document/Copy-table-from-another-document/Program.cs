using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

// Creates a new Word document.
using (WordDocument destinationDocument = new WordDocument())
{
    using (FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
    {
        // Opens an existing Word document.
        using (WordDocument sourceDocument = new WordDocument(inputStream, FormatType.Automatic))
        {
            // Add new section.
            WSection section = destinationDocument.AddSection() as WSection;
            // Get the table from source document and clone.
            WTable table = sourceDocument.LastSection.Tables[0].Clone() as WTable;
            // Insert the cloned table to destination document.
            section.Body.ChildEntities.Insert(0, table);
            // Saves the destination document.
            using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
            {
                destinationDocument.Save(outputStream, FormatType.Docx);
            }
        }
    }
}