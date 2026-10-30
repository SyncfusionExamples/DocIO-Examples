using SkiaSharp;
using Svg.Skia;
using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.IO;

using (WordDocument document = new WordDocument(@"../../../Data/Input.docx"))
{
    // Find all pictures by EntityType in the Word document.
    List<Entity> pictures = document.FindAllItemsByProperty(EntityType.Picture, null, null);
    for (int i = 0; i < pictures.Count; i++)
    {
        WPicture picture = pictures[i] as WPicture;
        byte[] imageBytes = picture.SvgData;

        // Store original properties
        float width = picture.Width;
        float height = picture.Height;
        if (imageBytes != null)
        {
            Byte[] svgBytes = picture.SvgData;
            var svg = new SKSvg();
            using var svgStream = new MemoryStream(svgBytes);
            svg.Load(svgStream);
            using var image = SKImage.FromPicture(
            svg.Picture,
            new SKSizeI(
            (int)svg.Picture.CullRect.Width,
            (int)svg.Picture.CullRect.Height));
            using var data = image.Encode(SKEncodedImageFormat.Png, 100);
            MemoryStream stream = new MemoryStream(data.ToArray());

            // Get parent paragraph
            WParagraph paragraph = picture.OwnerParagraph;

            // Find picture position in paragraph
            int index = paragraph.ChildEntities.IndexOf(picture);

            // Create replacement picture
            WPicture newPicture = new WPicture(picture.Document);
            newPicture.LoadImage(stream);

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