using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;
using Syncfusion.OfficeChartToImageConverter;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;

namespace Specific_range_of_pages_Word_to_image
{
    class Program
    {
        static void Main(string[] args)
        {
            //Load an existing Word document.
            using (WordDocument wordDocument = new WordDocument(Path.GetFullPath(@"../../Template.docx"), FormatType.Docx))
            {
                //Initialize the ChartToImageConverter for converting charts during Word to image conversion.
                wordDocument.ChartToImageConverter = new ChartToImageConverter();
                //Set the scaling mode for charts.
                wordDocument.ChartToImageConverter.ScalingMode = ScalingMode.Normal;
                //Convert a specific range of pages in Word document to images.
                Image[] images = wordDocument.RenderAsImages(1, 2, ImageType.Bitmap);
                int i = 0;
                foreach (Image image in images)
                {
                    //Save the images as jpeg.
                    image.Save(Path.GetFullPath(@"../../WordToImage_" + i + ".jpeg"), ImageFormat.Jpeg);
                    i++;
                }
            }
        }
    }
}
