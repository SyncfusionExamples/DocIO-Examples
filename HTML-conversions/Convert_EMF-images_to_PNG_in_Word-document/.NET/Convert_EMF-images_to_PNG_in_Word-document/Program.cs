using SkiaSharp;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using Syncfusion.Metafile;



using (FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Input.docx", FileMode.Open, FileAccess.ReadWrite))
{
    // Opens an input Word template.
    using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
    {
        // Find all picture in the document.
        List<Entity> images = document.FindAllItemsByProperty(EntityType.Picture, null, null);
        // Iterate through each image in the document.
        foreach (Entity entity in images)
        {
            WPicture picture = entity as WPicture;
            float width = picture.Width;
            float height = picture.Height;

            // Convert the image bytes into a memory stream.
            MemoryStream stream = new MemoryStream(picture.ImageBytes);
            // Create an Image object from the memory stream.
            Image image = Image.FromStream(stream);
            // Check if the image format is EMF.
            if (image.RawFormat.Equals(ImageFormat.Emf))
            {
                MemoryStream imageByteStream = new MemoryStream(picture.ImageBytes);
                //Create a new instance for the MetafileRenderer
                MetafileRenderer renderer = new MetafileRenderer();
                //Convert the Metafile to SKBitmap Image.
                SKBitmap skBitmap = renderer.ConvertToImage(imageByteStream);
                //Save the image as stream
                using (SKImage skImage = SKImage.FromBitmap(skBitmap))
                using (SKData data = skImage.Encode(SKEncodedImageFormat.Png, 100))

                using (MemoryStream pngStream = new MemoryStream())
                {
                    data.SaveTo(pngStream);
                    pngStream.Position = 0;
                    picture.LoadImage(pngStream);
                    picture.Height = height;
                    picture.Width = width;
                }
            }
        }
        // Save the modified document to the specified output file.
        using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Result.html"), FileMode.Create, FileAccess.Write))
        {
            document.Save(outputStream, FormatType.Html);
        }
    }
}


