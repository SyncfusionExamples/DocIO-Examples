using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

using (FileStream destinationStream = new FileStream(Path.GetFullPath(@"Data/DestinationDocument.docx"), FileMode.Open, FileAccess.ReadWrite))
{
    //Open the source document.
    using (WordDocument destinationDocument = new WordDocument(destinationStream, FormatType.Docx))
    {
        using (FileStream sourceStream = new FileStream(Path.GetFullPath(@"Data/SourceDocument.docx"), FileMode.Open, FileAccess.ReadWrite))
        {
            //Open the destination document.
            using (WordDocument sourceDocument = new WordDocument(sourceStream, FormatType.Docx))
            {
                // Retrieve the specific style from the source document.
                WParagraphStyle paraStyle = sourceDocument.Styles.FindByName("MyStyle") as WParagraphStyle;
                // Add the retrieved style to the destination document.
                destinationDocument.Styles.Add(paraStyle.Clone());
                // Get the first paragraph in the destination document.
                WParagraph paragraph = destinationDocument.Sections[0].Body.ChildEntities[0] as WParagraph;
                // Apply the retrieved style to the paragraph in the destination document.
                paragraph.ApplyStyle("MyStyle");
                // Save the destination document.
                using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    destinationDocument.Save(outputStream, FormatType.Docx);
                }
            }
        }
    }
}