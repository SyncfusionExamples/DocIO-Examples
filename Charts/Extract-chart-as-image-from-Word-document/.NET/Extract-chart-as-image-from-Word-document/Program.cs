using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using System.Collections.Generic;
using System.IO;

namespace Extract_Chart_As_Image_From_Word_Document
{
    class Program
    {
        static void Main(string[] args)
        {
            // Open the Word document file stream.
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                // Load the template Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    // Find all chart elements in the Word document.
                    List<Entity> charts = document.FindAllItemsByProperty(EntityType.Chart, null, null);

                    // Create an instance of DocIORenderer.
                    using (DocIORenderer renderer = new DocIORenderer())
                    {
                        // Iterate through each chart and convert it to an image.
                        for (int i = 0; i < charts.Count; i++)
                        {
                            WChart chart = charts[i] as WChart;

                            // Convert the chart to an image.
                            using (Stream stream = chart.SaveAsImage())
                            {
                                // Create an output image file stream with a unique filename.
                                using (FileStream fileStreamOutput = File.Create(Path.GetFullPath(@"Output/ChartToImage-" + i + ".jpeg")))
                                {
                                    // Copy the converted image stream into the output stream.
                                    stream.CopyTo(fileStreamOutput);
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
