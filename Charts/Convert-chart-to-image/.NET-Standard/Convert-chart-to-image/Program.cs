using System.IO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;

namespace Convert_chart_to_image
{
    class Program
    {
        static void Main(string[] args)
        {
			//Open the file as Stream.
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../TemplateWithChart.docx"), FileMode.Open))
            {
                //Load file stream into Word document.
                using (WordDocument wordDocument = new WordDocument(fileStream, Syncfusion.DocIO.FormatType.Automatic))
                {
                    //Get the first paragraph from the section. 
                    WParagraph paragraph = wordDocument.LastSection.Paragraphs[0];
                    //Get the chart element from the paragraph.
                    WChart chart = paragraph.ChildEntities[0] as WChart;
                    //Create an instance of DocIORenderer.
                    using (DocIORenderer renderer = new DocIORenderer())
                    {
                        //Convert chart to an image.
                        using (Stream stream = chart.SaveAsImage())
                        {
                            //Create the output image file stream. 
                            using (FileStream fileStreamOutput = File.Create(Path.GetFullPath(@"../../../ChartToImage.jpeg")))
                            {
                                //Copies the converted image stream into created output stream.
                                stream.CopyTo(fileStreamOutput);
                            }
                        }
                    }
                }
            }
        }
    }
}
