using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;

namespace Create_bubble_chart
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
                chart.ChartType = OfficeChartType.Bubble;
                //Set chart data.
                chart.ChartData.SetValue(1, 1, "X-Values");
                chart.ChartData.SetValue(2, 1, -10);
                chart.ChartData.SetValue(3, 1, -20);
                chart.ChartData.SetValue(4, 1, -30);
                chart.ChartData.SetValue(5, 1, 10);
                chart.ChartData.SetValue(6, 1, 20);
                chart.ChartData.SetValue(7, 1, 30);
                chart.ChartData.SetValue(1, 2, "Y-Values");
                chart.ChartData.SetValue(2, 2, -100);
                chart.ChartData.SetValue(3, 2, -200);
                chart.ChartData.SetValue(4, 2, -300);
                chart.ChartData.SetValue(5, 2, 100);
                chart.ChartData.SetValue(6, 2, 200);
                chart.ChartData.SetValue(7, 2, 300);
                chart.ChartData.SetValue(1, 3, "Size");
                chart.ChartData.SetValue(2, 3, 1);
                chart.ChartData.SetValue(3, 3, -1);
                chart.ChartData.SetValue(4, 3, 1);
                chart.ChartData.SetValue(5, 3, -1);
                chart.ChartData.SetValue(6, 3, 1);
                chart.ChartData.SetValue(7, 3, -1);
                //Set a Chart Title.
                chart.ChartTitle = "Bubble Chart";
                //Set Datalabels.
                IOfficeChartSerie series = chart.Series.Add();
                //Set the data range of chart series – start row, start column, end row, and end column.
                series.CategoryLabels = chart.ChartData[2, 1, 7, 1];
                series.Values = chart.ChartData[2, 2, 7, 2];
                series.Bubbles = chart.ChartData[2, 3, 7, 3];
                series.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
                series.DataPoints.DefaultDataPoint.DataLabels.Position = OfficeDataLabelPosition.Center;
                //Set legend.
                chart.HasLegend = false;
                //Create a file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Sample.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Save the Word document to the file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
