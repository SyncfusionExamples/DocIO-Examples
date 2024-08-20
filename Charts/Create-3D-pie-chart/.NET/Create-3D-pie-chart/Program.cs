using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;
using System.IO;

namespace Create_3D_pie_chart
{
    class Program
    {
        static void Main(string[] args)
        {
            //Create a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Add a section to the document.
                IWSection section = document.AddSection();
                //Add a paragraph to the section.
                IWParagraph paragraph = section.AddParagraph();
                //Create and append the chart to the paragraph.
                WChart chart = paragraph.AppendChart(446, 270);
                //Set chart type.
                chart.ChartType = OfficeChartType.Pie_3D;
                //Set chart data.
                chart.ChartData.SetValue(2, 1, "Food");
                chart.ChartData.SetValue(3, 1, "Fruits");
                chart.ChartData.SetValue(4, 1, "Vegetables");
                chart.ChartData.SetValue(5, 1, "Dairy");
                chart.ChartData.SetValue(6, 1, "Protein");
                chart.ChartData.SetValue(7, 1, "Grains");
                chart.ChartData.SetValue(2, 2, "Percentage");
                chart.ChartData.SetValue(3, 2, 36);
                chart.ChartData.SetValue(4, 2, 14);
                chart.ChartData.SetValue(5, 2, 13);
                chart.ChartData.SetValue(6, 2, 28);
                chart.ChartData.SetValue(7, 2, 9);
                //Set region of Chart data.
                chart.DataRange = chart.ChartData[2, 1, 7, 2];
                //Set chart series in the column for assigned data region.
                chart.IsSeriesInRows = false;
                //Set a Chart Title.
                chart.ChartTitle = "3D Pie Chart";
                //Set Datalabels.
                IOfficeChartSerie serie = chart.Series[0];

                serie.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
                //Set legend.
                chart.HasLegend = true;
                chart.Legend.Position = OfficeLegendPosition.Bottom;
                //Set elevation.
                chart.Elevation = 30;
                //Create a file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Save the Word document to the file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }              
    }    
}
