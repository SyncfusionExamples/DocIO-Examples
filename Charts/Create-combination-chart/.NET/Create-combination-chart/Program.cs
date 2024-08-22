using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;

namespace Create_combination_chart
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
                chart.ChartData.SetValue(1, 1, "Month");
                chart.ChartData.SetValue(2, 1, "Jan");
                chart.ChartData.SetValue(3, 1, "Feb");
                chart.ChartData.SetValue(4, 1, "Mar");
                chart.ChartData.SetValue(5, 1, "Apr");
                chart.ChartData.SetValue(6, 1, "May");
                chart.ChartData.SetValue(7, 1, "Jun");
                chart.ChartData.SetValue(8, 1, "Jul");
                chart.ChartData.SetValue(9, 1, "Aug");
                chart.ChartData.SetValue(10, 1, "Sep");
                chart.ChartData.SetValue(11, 1, "Oct");
                chart.ChartData.SetValue(12, 1, "Nov");
                chart.ChartData.SetValue(13, 1, "Dec");
                chart.ChartData.SetValue(1, 2, "Rainy Days");
                chart.ChartData.SetValue(2, 2, 12);
                chart.ChartData.SetValue(3, 2, 11);
                chart.ChartData.SetValue(4, 2, 10);
                chart.ChartData.SetValue(5, 2, 9);
                chart.ChartData.SetValue(6, 2, 8);
                chart.ChartData.SetValue(7, 2, 6);
                chart.ChartData.SetValue(8, 2, 4);
                chart.ChartData.SetValue(9, 2, 6);
                chart.ChartData.SetValue(10, 2, 7);
                chart.ChartData.SetValue(11, 2, 8);
                chart.ChartData.SetValue(12, 2, 10);
                chart.ChartData.SetValue(13, 2, 11);
                chart.ChartData.SetValue(1, 3, "Profit");
                chart.ChartData.SetValue(2, 3, 3574);
                chart.ChartData.SetValue(3, 3, 4708);
                chart.ChartData.SetValue(4, 3, 5332);
                chart.ChartData.SetValue(5, 3, 6693);
                chart.ChartData.SetValue(6, 3, 8843);
                chart.ChartData.SetValue(7, 3, 12347);
                chart.ChartData.SetValue(8, 3, 15180);
                chart.ChartData.SetValue(9, 3, 11198);
                chart.ChartData.SetValue(10, 3, 9739);
                chart.ChartData.SetValue(11, 3, 9846);
                chart.ChartData.SetValue(12, 3, 6620);
                chart.ChartData.SetValue(13, 3, 5085);
                //Set region of Chart data.
                chart.DataRange = chart.ChartData[1, 1, 13, 3];
                //Set chart series in the column for assigned data region.
                chart.IsSeriesInRows = false;
                //Set a Chart Title.
                chart.ChartTitle = "Combination Chart";
                //Set Datalabels.
                IOfficeChartSerie series1 = chart.Series[0];
                IOfficeChartSerie series2 = chart.Series[1];
                //Set Serie type.
                series1.SerieType = OfficeChartType.Column_Clustered;
                series2.SerieType = OfficeChartType.Line;
                series2.UsePrimaryAxis = false;
                //Set Datalabels.
                series1.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
                //Set legend.
                chart.HasLegend = true;
                chart.Legend.Position = OfficeLegendPosition.Bottom;
                //Set chart type.
                chart.ChartType = OfficeChartType.Combination_Chart;
                //Set secondary axis on right side.
                chart.SecondaryValueAxis.TickLabelPosition = OfficeTickLabelPosition.TickLabelPosition_High;
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
