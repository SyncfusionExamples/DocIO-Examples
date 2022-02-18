using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;
using Syncfusion.OfficeChartToImageConverter;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace Convert_chart_to_image
{
    class Program
    {
        static void Main(string[] args)
        {
            //Loads an existing Word document.
            using (WordDocument wordDocument = new WordDocument(Path.GetFullPath(@"../../TemplateWithChart.docx"), FormatType.Automatic))
            {
                //Initializes the ChartToImageConverter for converting charts during Word to image conversion.
                wordDocument.ChartToImageConverter = new ChartToImageConverter();
                //Sets the scaling mode for charts. (Normal mode reduces the file size)
                wordDocument.ChartToImageConverter.ScalingMode = ScalingMode.Normal;
                //Gets the first paragraph from section.
                WParagraph paragraph = wordDocument.LastSection.Paragraphs[0];
                //Gets the chart element in the paragarph item.
                WChart chart = paragraph.ChildEntities[0] as WChart;
                //Creating the memory stream for chart image.
                using (MemoryStream stream = new MemoryStream())
                {
                    //Converts chart to image.
                    wordDocument.ChartToImageConverter.SaveAsImage(chart.OfficeChart, stream);
                    Image image = Image.FromStream(stream);
                    //Saving image stream to file.
                    image.Save(@"../../ChartToImage.jpeg", ImageFormat.Jpeg);
                }
            }
        }
    }
}
