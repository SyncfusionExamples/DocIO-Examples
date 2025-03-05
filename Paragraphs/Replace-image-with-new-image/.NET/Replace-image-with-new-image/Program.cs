using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Replace_image_with_new_image
{
    class Program
    {
        static void Main(string[] args)
        {

            // Open the input Word document for reading and writing
            using (FileStream inputStream = new FileStream(@"Data/Template.docx", FileMode.Open, FileAccess.ReadWrite))
            {
                WordDocument document = new WordDocument(inputStream, FormatType.Docx);
                // Find and process all images in the document
                foreach (Entity entity in document.FindAllItemsByProperty(EntityType.Picture, null, null))
                {
                    WPicture picture = (WPicture)entity;
                    float width = picture.Width, height = picture.Height;
                    // Check if the image is not an SVG
                    if (picture.SvgData == null)
                    {
                        // Replace the existing image with a new one
                        using (FileStream imageStream = new FileStream(@"Data/Picture.png", FileMode.Open))
                        {
                            picture.LoadImage(imageStream);
                        }
                        // Preserve the original dimensions
                        picture.LockAspectRatio = false;
                        picture.Width = width;
                        picture.Height = height;
                        picture.LockAspectRatio = true;
                    }
                    else
                    {
                        // Handle SVG conversion to raster image
                        WParagraph ownerParagraph = picture.OwnerParagraph;
                        int index = ownerParagraph.ChildEntities.IndexOf(picture);
                        // Remove the existing SVG image
                        ownerParagraph.ChildEntities.Remove(picture);
                        // Create a new image and insert it in the same place
                        WPicture newPicture = new WPicture(document);
                        using (FileStream imageStream = new FileStream(@"Data/Picture.png", FileMode.Open))
                        {
                            newPicture.LoadImage(imageStream);
                        }
                        // Set the same dimensions as the original SVG
                        newPicture.Width = width;
                        newPicture.Height = height;
                        ownerParagraph.ChildEntities.Insert(index, newPicture);
                    }
                }
                // Save the modified document
                using (FileStream outputStream = new FileStream(@"Output/Result.docx", FileMode.Create, FileAccess.Write))
                {
                    document.Save(outputStream, FormatType.Docx);
                }
            }
        }
    }
}
