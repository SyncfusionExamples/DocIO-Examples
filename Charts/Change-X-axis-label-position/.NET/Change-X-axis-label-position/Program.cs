using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;

namespace Change_X_axis_label_position
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/Input.docx"), FileMode.Open))
            {
                //Load an existing Word document.
                using (WordDocument wordDocument = new WordDocument(fileStream, Syncfusion.DocIO.FormatType.Automatic))
                {
                    // Get the paragraph.
                    WParagraph paragraph = wordDocument.LastParagraph;
                    //Get the chart entity.
                    WChart chart = paragraph.ChildEntities[0] as WChart;
                    //Set X-axis label position to the bottom of the chart.
                    chart.PrimaryCategoryAxis.TickLabelPosition = OfficeTickLabelPosition.TickLabelPosition_Low;
                    //Create a file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Save the Word document to the file stream.
                        wordDocument.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }
    }
}
