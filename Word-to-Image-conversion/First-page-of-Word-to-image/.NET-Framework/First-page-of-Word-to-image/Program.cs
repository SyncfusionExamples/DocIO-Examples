using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;
using Syncfusion.OfficeChartToImageConverter;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;

namespace First_page_of_Word_to_image
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
                //Convert the first page of the Word document into an image.
                Image image = wordDocument.RenderAsImages(0, ImageType.Bitmap);
                //Save the image as jpeg.
                image.Save(Path.GetFullPath(@"../../WordToImage.jpeg"), ImageFormat.Jpeg);      
            }
        }
    }
}
