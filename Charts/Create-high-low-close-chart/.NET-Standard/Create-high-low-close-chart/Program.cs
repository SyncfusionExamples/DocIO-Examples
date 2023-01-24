using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;

namespace Create_high_low_close_chart
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
                //Set chart data.
                chart.ChartData.SetValue(1, 1, "Date");
                chart.ChartData.SetValue(2, 1, "High");
                chart.ChartData.SetValue(3, 1, "Low");
                chart.ChartData.SetValue(4, 1, "Close");
                chart.ChartData.SetValue(1, 2, "1-Apr-17");
                chart.ChartData.SetValue(2, 2, 50);
                chart.ChartData.SetValue(3, 2, 10);
                chart.ChartData.SetValue(4, 2, 40);
                chart.ChartData.SetValue(1, 3, "2-Apr-17");
                chart.ChartData.SetValue(2, 3, 60);
                chart.ChartData.SetValue(3, 3, 20);
                chart.ChartData.SetValue(4, 3, 30);
                chart.ChartData.SetValue(1, 4, "3-Apr-17");
                chart.ChartData.SetValue(2, 4, 55);
                chart.ChartData.SetValue(3, 4, 15);
                chart.ChartData.SetValue(4, 4, 45);
                chart.ChartData.SetValue(1, 5, "4-Apr-17");
                chart.ChartData.SetValue(2, 5, 65);
                chart.ChartData.SetValue(3, 5, 25);
                chart.ChartData.SetValue(4, 5, 35);
                chart.ChartData.SetValue(1, 6, "5-Apr-17");
                chart.ChartData.SetValue(2, 6, 70);
                chart.ChartData.SetValue(3, 6, 30);
                chart.ChartData.SetValue(4, 6, 60);
                //Set region of Chart data.
                chart.DataRange = chart.ChartData[1, 1, 4, 6];
                //Set chart series in the column for assigned data region.
                chart.IsSeriesInRows = true;
                //Set chart type.
                chart.ChartType = OfficeChartType.Stock_HighLowClose;
                //Set a Chart Title.
                chart.ChartTitle = "High-Low-Close Chart";
                //Set primary category axis.
                chart.PrimaryCategoryAxis.NumberFormat = "dd-MMM-yy";
                //Set Datalabels.
                IOfficeChartSerie series1 = chart.Series[0];
                IOfficeChartSerie series2 = chart.Series[1];
                IOfficeChartSerie series3 = chart.Series[2];

                series1.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
                series1.DataPoints.DefaultDataPoint.DataLabels.IsSeriesName = true;
                series1.SerieFormat.MarkerStyle = OfficeChartMarkerType.Circle;
                series1.SerieFormat.MarkerForegroundColorIndex = OfficeKnownColors.LightGreen;

                series2.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
                series2.DataPoints.DefaultDataPoint.DataLabels.IsSeriesName = true;
                series2.SerieFormat.MarkerStyle = OfficeChartMarkerType.Circle;
                series2.SerieFormat.MarkerForegroundColorIndex = OfficeKnownColors.Red;

                series3.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
                series3.DataPoints.DefaultDataPoint.DataLabels.IsSeriesName = true;
                series3.SerieFormat.MarkerStyle = OfficeChartMarkerType.Circle;
                series3.SerieFormat.MarkerForegroundColorIndex = OfficeKnownColors.Light_yellow;
                //Set legend.
                chart.HasLegend = true;
                chart.Legend.Position = OfficeLegendPosition.Bottom;
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
