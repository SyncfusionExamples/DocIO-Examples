using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    // Opens the template Word document.
    using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic))
    {
        // Variables to store the original width and height of each picture.
        float pictureWidth, pictureHeight = 0;
        // Find the picture with alternative text.
        WPicture picture = document.FindItemByProperty(EntityType.Picture, "AlternativeText", "Adventure") as WPicture;
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
        using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
        {
            // Saves the modified Word document.
            document.Save(outputFileStream, FormatType.Docx);
        }
    }
}
