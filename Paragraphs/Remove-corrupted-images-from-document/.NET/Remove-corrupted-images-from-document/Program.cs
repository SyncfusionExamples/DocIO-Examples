using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.Collections.Generic;
using System.IO;

namespace Remove_corrupted_images_from_document
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream inputStream = new FileStream(@"Data/Template.docx", FileMode.Open, FileAccess.ReadWrite))
            {
                // Opens an input Word template.
                WordDocument document = new WordDocument(inputStream, FormatType.Docx);
                // Find all picture in the document.
                List<Entity> images = document.FindAllItemsByProperty(EntityType.Picture, null, null);
                // Iterate through each image in the document.
                foreach (Entity entity in images)
                {
                    WPicture picture = entity as WPicture;
                    // Convert the image bytes into a memory stream.
                    MemoryStream stream = new MemoryStream(picture.ImageBytes);
                    // Create an Image object from the memory stream.
                    Image image = Image.FromStream(stream);
                    // Check if the image format is unknown (corrupt or unsupported image).
                    if (image.RawFormat == ImageFormat.Unknown)
                    {
                        // Remove the invalid image from the document.
                        picture.OwnerParagraph.ChildEntities.Remove(picture);
                    }
                }
                // Save the modified document to the specified output file.
                using (FileStream outputStream = new FileStream(@"../../../Output/Result.docx", FileMode.Create, FileAccess.Write))
                {
                    document.Save(outputStream, FormatType.Docx);
                }
            }
        }
    }
}
