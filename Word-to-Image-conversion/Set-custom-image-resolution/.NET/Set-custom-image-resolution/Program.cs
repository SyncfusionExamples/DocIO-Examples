using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using System.Drawing;
using System.Drawing.Imaging;


// Open the input Word document stream in read mode
using (FileStream docStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read))
{
    // Load the Word document using Syncfusion DocIO
    using (WordDocument document = new WordDocument(docStream, FormatType.Automatic))
    {
        // Create an instance of DocIORenderer to render the Word document as images
        using (DocIORenderer render = new DocIORenderer())
        {
            // Convert all pages of the Word document to image streams
            Stream[] imageStreams = document.RenderAsImages();

            // Iterate through each image stream (one per page)
            for (int i = 0; i < imageStreams.Length; i++)
            {
                // Reset the stream position to the beginning
                imageStreams[i].Position = 0;

                // Define custom dimensions for the output image
                int customWidth = 1500;
                int customHeight = 1500;

                // Load the image from stream
                Image image = Image.FromStream(imageStreams[i]);

                // Save the image to a new memory stream in PNG format
                MemoryStream stream = new MemoryStream();
                image.Save(stream, ImageFormat.Png);

                // Create a new bitmap with custom size and pixel format
                Bitmap bitmap = new Bitmap(customWidth, customHeight, PixelFormat.Format32bppPArgb);

                // Create graphics object to draw on the bitmap
                Graphics graphics = Graphics.FromImage(bitmap);

                // Set bitmap resolution to 300 DPI
                bitmap.SetResolution(300, 300);

                // Draw the resized image onto the custom-sized bitmap
                graphics.DrawImage(Image.FromStream(stream), new Rectangle(0, 0, bitmap.Width, bitmap.Height));

                // Save the final bitmap image to output folder
                bitmap.Save(Path.GetFullPath(@"Output/Image_" + i + ".png"));
            }
        }
    }
}
