using Syncfusion.DocIO.DLS;
using System.Collections.Generic;
using System.IO;
using System.Drawing.Imaging;


using (WordDocument document = new WordDocument(@"../../../Data/Sample.docx"))
{
    // Find all pictures by EntityType in the Word document.
    List<Entity> pictures = document.FindAllItemsByProperty(EntityType.Picture, null, null);
    for (int i = 0; i < pictures.Count; i++)
    {
        WPicture picture = pictures[i] as WPicture;
        if (IsTiff(picture))
        {
            // Store original properties
            float width = picture.Width;
            float height = picture.Height;
            byte[] imageBytes = picture.ImageBytes;

            if (imageBytes == null || imageBytes.Length == 0)
                continue;

            using MemoryStream tiffStream = new MemoryStream(imageBytes);
            using System.Drawing.Image image = System.Drawing.Image.FromStream(tiffStream);

            MemoryStream jpegStream = new MemoryStream();

            // Save as JPEG
            image.Save(jpegStream, ImageFormat.Jpeg);

            jpegStream.Position = 0;

            // Get parent paragraph
            WParagraph paragraph = picture.OwnerParagraph;

            // Find picture position in paragraph
            int index = paragraph.ChildEntities.IndexOf(picture);

            // Create replacement picture
            WPicture newPicture = new WPicture(picture.Document);
            newPicture.LoadImage(jpegStream);

            // Preserve dimensions
            newPicture.Width = width;
            newPicture.Height = height;

            // Insert new picture at same location
            paragraph.ChildEntities.Insert(index, newPicture);

            // Remove original SVG picture
            paragraph.ChildEntities.Remove(picture);
        }
    }
    document.Save(@"../../../Output/Output.docx");
}

/// <summary>
/// Determines whether a picture contains a TIFF image.
/// </summary>
/// <param name="picture">The picture.</param>
static bool IsTiff(WPicture picture)
{
    byte[] imageBytes = picture.ImageBytes;

    if (imageBytes == null || imageBytes.Length < 4)
        return false;

    // Little-endian TIFF header: II*<NUL>
    bool littleEndianTiff =
        imageBytes[0] == 0x49 &&
        imageBytes[1] == 0x49 &&
        imageBytes[2] == 0x2A &&
        imageBytes[3] == 0x00;

    // Big-endian TIFF header: MM<NUL>*
    bool bigEndianTiff =
        imageBytes[0] == 0x4D &&
        imageBytes[1] == 0x4D &&
        imageBytes[2] == 0x00 &&
        imageBytes[3] == 0x2A;

    return littleEndianTiff || bigEndianTiff;
}

