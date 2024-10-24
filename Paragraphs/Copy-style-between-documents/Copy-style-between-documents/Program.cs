using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

using (FileStream sourceStream = new FileStream(Path.GetFullPath(@"Data/SourceDocument.docx"), FileMode.Open, FileAccess.ReadWrite))
{
    //Open the source document.
    using (WordDocument sourceDocument = new WordDocument(sourceStream, FormatType.Docx))
    {
        using (FileStream destinationStream = new FileStream(Path.GetFullPath(@"Data/DestinationDocument.docx"), FileMode.Open, FileAccess.ReadWrite))
        {
            //Open the destination document.
            using (WordDocument destinationDocument = new WordDocument(destinationStream, FormatType.Docx))
            {
                //Get the style from source document.
                WParagraphStyle paraStyle = sourceDocument.Styles.FindByName("User_defined_style") as WParagraphStyle;
                //Add it into destination document.
                destinationDocument.Styles.Add(paraStyle.Clone());
                //Get the first paragraph in destination document.
                WParagraph paragraph = destinationDocument.Sections[0].Body.ChildEntities[0] as WParagraph;
                //Applies the new style to paragraph.
                paragraph.ApplyStyle("User_defined_style");
                //Saves the destination document.
                using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    destinationDocument.Save(outputStream, FormatType.Docx);
                }
            }
        }
    }
}