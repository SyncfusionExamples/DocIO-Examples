using System;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;
using Syncfusion.OfficeChartToImageConverter;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;

namespace Custom_image_resolution
{
    class Program
    {
        static void Main(string[] args)
        {
            //Loads an existing Word document.
            using (WordDocument wordDocument = new WordDocument(Path.GetFullPath(@"../../Template.docx"), FormatType.Docx))
            {
                //Initialize the ChartToImageConverter for converting charts during Word to image conversion.
                wordDocument.ChartToImageConverter = new ChartToImageConverter();
                //Set the scaling mode for charts (Normal mode reduces the file size).
                wordDocument.ChartToImageConverter.ScalingMode = ScalingMode.Normal;
                //Convert the word document to images.
                Image[] images = wordDocument.RenderAsImages(ImageType.Metafile);
                //Declare the variables to hold custom width and height.
                int customWidth = 1500;
                int customHeight = 1500;
                foreach (Image image in images)
                {
                    MemoryStream stream = new MemoryStream();
                    image.Save(stream, ImageFormat.Png);
                    //Create a bitmap of specific width and height.
                    Bitmap bitmap = new Bitmap(customWidth, customHeight, PixelFormat.Format32bppPArgb);
                    //Get the graphics from an image.
                    Graphics graphics = Graphics.FromImage(bitmap);
                    //Set the resolution.
                    bitmap.SetResolution(300, 300);
                    //Recreate the image in custom size.
                    graphics.DrawImage(System.Drawing.Image.FromStream(stream), new Rectangle(0, 0, bitmap.Width, bitmap.Height));
                    //Save the image as a bitmap.
                    bitmap.Save(Path.GetFullPath(@"../../ImageOutput" + Guid.NewGuid().ToString() + ".png"));
                }
            }
        }
    }
}
