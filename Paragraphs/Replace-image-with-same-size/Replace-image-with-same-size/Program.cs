using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    // Opens the template Word document.
    using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
    {
        // Variables to store the original width and height of each picture.
        float pictureWidth, pictureHeight = 0;
        // Retrieve the body of the first section.
        WTextBody textbody = document.Sections[0].Body;
        // Iterates through each paragraph in the document's body.
        foreach (WParagraph paragraph in textbody.Paragraphs)
        {
            // Iterates through each item within the paragraph.
            foreach (ParagraphItem item in paragraph.ChildEntities)
            {
                // Checks if the item is a picture.
                if (item is WPicture)
                {
                    // Casts the item to WPicture to access picture properties.
                    WPicture picture = item as WPicture;
                    // Store the original width and height of the picture.
                    pictureWidth = picture.Width;
                    pictureHeight = picture.Height;
                    // Open the new image file to replace the existing picture.
                    FileStream imageStream = new FileStream(Path.GetFullPath(@"Data/Image.png"), FileMode.Open, FileAccess.ReadWrite);
                    // Load the new image into the existing picture element.
                    picture.LoadImage(imageStream);
                    // Restore the original width and height to maintain the picture's dimensions.
                    picture.Width = pictureWidth;
                    picture.Height = pictureHeight;
                }
            }
        }
        using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
        {
            // Saves the modified Word document.
            document.Save(outputFileStream, FormatType.Docx);
        }
    }
}
