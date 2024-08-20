using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;

namespace Create_wireframe_3D_surface_chart
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Add a section to the document.
                IWSection section = document.AddSection();
                //Add a paragraph to the section.
                IWParagraph paragraph = section.AddParagraph();
                //Create and append the chart to the paragraph.
                WChart chart = paragraph.AppendChart(446, 270);
                //Set chart data.
                chart.ChartData.SetValue(1, 1, "Course");
                chart.ChartData.SetValue(2, 1, "English");
                chart.ChartData.SetValue(3, 1, "Physics");
                chart.ChartData.SetValue(4, 1, "Maths");
                chart.ChartData.SetValue(5, 1, "History");
                chart.ChartData.SetValue(6, 1, "Language 1");
                chart.ChartData.SetValue(7, 1, "Language 2");
                chart.ChartData.SetValue(1, 2, "SchoolA");
                chart.ChartData.SetValue(2, 2, 63);
                chart.ChartData.SetValue(3, 2, 61);
                chart.ChartData.SetValue(4, 2, 62);
                chart.ChartData.SetValue(5, 2, 46);
                chart.ChartData.SetValue(6, 2, 60);
                chart.ChartData.SetValue(7, 2, 63);
                chart.ChartData.SetValue(1, 3, "SchoolB");
                chart.ChartData.SetValue(2, 3, 53);
                chart.ChartData.SetValue(3, 3, 55);
                chart.ChartData.SetValue(4, 3, 51);
                chart.ChartData.SetValue(5, 3, 53);
                chart.ChartData.SetValue(6, 3, 56);
                chart.ChartData.SetValue(7, 3, 58);
                chart.ChartData.SetValue(1, 4, "SchoolC");
                chart.ChartData.SetValue(2, 4, 45);
                chart.ChartData.SetValue(3, 4, 65);
                chart.ChartData.SetValue(4, 4, 64);
                chart.ChartData.SetValue(5, 4, 66);
                chart.ChartData.SetValue(6, 4, 64);
                chart.ChartData.SetValue(7, 4, 64);
                //Set region of Chart data.
                chart.DataRange = chart.ChartData[1, 1, 7, 4];
                //Set chart type.
                chart.ChartType = OfficeChartType.Surface_NoColor_3D;
                //Set chart series in the column for assigned data region.
                chart.IsSeriesInRows = false;
                //Set a Chart Title.
                chart.ChartTitle = "Wireframe 3D Surface Chart";
                //Set legend.
                chart.HasLegend = true;
                chart.Legend.Position = OfficeLegendPosition.Bottom;
                //Set elevation, rotation, and perspective.
                chart.Rotation = 20;
                chart.Elevation = 20;
                chart.Perspective = 20;
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
