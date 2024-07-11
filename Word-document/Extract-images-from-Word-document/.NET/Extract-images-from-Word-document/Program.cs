using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.IO;

namespace Extract_images_from_Word_document
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Open the file as a stream.
            using (FileStream docStream = new FileStream(Path.GetFullPath(@"../../../Data/Template.docx"), FileMode.Open, FileAccess.Read))
            {
                // Load the file stream into a Word document.
                using (WordDocument document = new WordDocument(docStream, FormatType.Docx))
                {
                    // Find all pictures by EntityType in the Word document.
                    List<Entity> pictures = document.FindAllItemsByProperty(EntityType.Picture, null, null);

                    // Iterate through the pictures and save each one as an image file.
                    for (int i = 0; i < pictures.Count; i++)
                    {
                        WPicture image = pictures[i] as WPicture;

                        // Use a MemoryStream to handle the image bytes from the picture.
                        using (MemoryStream memoryStream = new MemoryStream(image.ImageBytes))
                        {
                            // Define the path where the image will be saved.
                            string imagePath = Path.GetFullPath(@"../../../Image-" + i + ".jpeg");

                            // Create a FileStream to write the image to the specified path.
                            using (FileStream filestream = new FileStream(imagePath, FileMode.Create, FileAccess.Write))
                            {
                                memoryStream.CopyTo(filestream);
                            }
                        }
                    }
                }
            }
        }
    }
}
