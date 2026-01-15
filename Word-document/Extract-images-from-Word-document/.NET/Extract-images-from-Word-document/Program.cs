using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
using System;
using System.Collections.Generic;
using System.IO;

namespace Extract_images_from_Word_document
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Open the Word document.
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"Data/Template.docx")))
            {
                // Find all pictures by EntityType in the Word document.
                List<Entity> pictures = document.FindAllItemsByProperty(EntityType.Picture, null, null);
                //To save unique images with identifier
                int count = 0;
                // Iterate through the pictures and save each one as an image file.
                for (int i = 0; i < pictures.Count; i++)
                {
                    WPicture image = pictures[i] as WPicture;
                    ExtractImages(image.ImageBytes, count);
                    count++;
                }
                // Find all smart arts by EntityType in the Word document.
                List<Entity> smartArts = document.FindAllItemsByProperty(EntityType.SmartArt, null, null);
                // Iterate through the smart art.
                for (int i = 0; i < smartArts.Count; i++)
                {
                    WSmartArt smartArt = smartArts[i] as WSmartArt;
                    //Extract background image in the smart art
                    if (smartArt.Background.PictureFill.ImageBytes != null)
                    {
                        ExtractImages(smartArt.Background.PictureFill.ImageBytes, count);
                        count++;
                    }
                    //Traverse through all nodes inside the SmartArt.
                    foreach (IOfficeSmartArtNode node in smartArt.Nodes)
                    {
                        foreach (IOfficeSmartArtShape shape in node.Shapes)
                        {
                            //If shape fill type is picture, then extract the image.
                            if (shape.Fill.FillType == OfficeShapeFillType.Picture)
                            {
                                ExtractImages(shape.Fill.PictureFill.ImageBytes, count);
                                count++;
                            }
                        }
                    }
                }
            }
        }
        /// <summary>
        /// Extracts image data from a byte array and saves it as a JPEG file with a unique identifier.
        /// </summary>
        /// <param name="imageBytes">The byte array containing image data to be saved</param>
        /// <param name="count">A unique identifier used to name the output image file.</param>
        private static void ExtractImages(byte[] imageBytes, int count)
        {
            // Use a MemoryStream to handle the image bytes from the picture.
            using (MemoryStream memoryStream = new MemoryStream(imageBytes))
            {
                // Define the path where the image will be saved.
                string imagePath = Path.GetFullPath(@"../../../Output/Image-" + count + ".jpeg");
                // Create a FileStream to write the image to the specified path.
                using (FileStream filestream = new FileStream(imagePath, FileMode.Create, FileAccess.Write))
                {
                    memoryStream.CopyTo(filestream);
                }
            }
        }
    }
}
