using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using System.Collections.Generic;
using System.IO;

namespace Convert_chart_to_image_and_replace
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the file as Stream.
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open))
            {
                //Load file stream into Word document.
                using (WordDocument wordDocument = new WordDocument(fileStream, FormatType.Docx))
                {
                    List<Entity> charts = wordDocument.FindAllItemsByProperty(EntityType.Chart, null, null);
                    for (int i=0; i<charts.Count; i++)
                    {
                        WChart chart = (WChart)charts[i];
                        //Get owner paragraph of chart.
                        WParagraph paragraph = chart.OwnerParagraph;
                        //Get index of the chart in the paragraph.
                        int chartIndex = paragraph.ChildEntities.IndexOf(chart);
                        //Create an instance of DocIORenderer.
                        using (DocIORenderer renderer = new DocIORenderer())
                        {
                            //Convert chart to an image.
                            using (Stream stream = chart.SaveAsImage())
                            {
                                //Create an instance of WPicture.
                                WPicture picture = new WPicture(wordDocument);
                                //Load image from stream.
                                picture.LoadImage((stream as MemoryStream).ToArray());
                                //Set width and height of the image.
                                picture.Width = chart.Width;
                                picture.Height = chart.Height;
                                //Add image to the paragraph at chart index.
                                if (chartIndex < paragraph.ChildEntities.Count)
                                    paragraph.ChildEntities.Insert(chartIndex, picture);
                                //Remove chart from the paragraph.
                                paragraph.ChildEntities.Remove(chart);
                            }
                        }
                    }

                    //Create a file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to the file stream.
                        wordDocument.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
