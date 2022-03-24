using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;
using Syncfusion.OfficeChartToImageConverter;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;

namespace Convert_Word_to_image
{
    class Program
    {
        static void Main(string[] args)
        {
            //Loads an existing Word document.
            using (WordDocument wordDocument = new WordDocument(Path.GetFullPath(@"../../Template.docx"), FormatType.Docx))
            {
                //Initializes the ChartToImageConverter for converting charts during Word to image conversion.
                wordDocument.ChartToImageConverter = new ChartToImageConverter();
                //Sets the scaling mode for charts (Normal mode reduces the file size).
                wordDocument.ChartToImageConverter.ScalingMode = ScalingMode.Normal;
                //Converts word document to image.
                Image[] images = wordDocument.RenderAsImages(ImageType.Bitmap);
                int i = 0;
                foreach (Image image in images)
                {
                    //Saves the images as jpeg.
                    image.Save(Path.GetFullPath(@"../../WordToImage_" + i + ".jpeg"), ImageFormat.Jpeg);
                    i++;
                }
            }
        }
    }
}
