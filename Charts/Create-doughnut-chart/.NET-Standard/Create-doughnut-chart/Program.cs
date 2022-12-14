using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;

namespace Create_doughnut_chart
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
                chart.ChartType = OfficeChartType.Doughnut;
                //Set chart data.
                chart.ChartData.SetValue(1, 1, "Company");
                chart.ChartData.SetValue(2, 1, "Company A");
                chart.ChartData.SetValue(3, 1, "Company B");
                chart.ChartData.SetValue(4, 1, "Company C");
                chart.ChartData.SetValue(5, 1, "Company D");
                chart.ChartData.SetValue(6, 1, "Company E");
                chart.ChartData.SetValue(7, 1, "Others");
                chart.ChartData.SetValue(1, 2, "2016");
                chart.ChartData.SetValue(2, 2, 28);
                chart.ChartData.SetValue(3, 2, 5);
                chart.ChartData.SetValue(4, 2, 17);
                chart.ChartData.SetValue(5, 2, 18);
                chart.ChartData.SetValue(6, 2, 17);
                chart.ChartData.SetValue(7, 2, 15);
                chart.ChartData.SetValue(1, 3, "2017");
                chart.ChartData.SetValue(2, 3, 25);
                chart.ChartData.SetValue(3, 3, 9);
                chart.ChartData.SetValue(4, 3, 19);
                chart.ChartData.SetValue(5, 3, 22);
                chart.ChartData.SetValue(6, 3, 15);
                chart.ChartData.SetValue(7, 3, 10);
                //Set region of Chart data.
                chart.DataRange = chart.ChartData[1, 1, 7, 3];
                //Set chart series in the column for assigned data region.
                chart.IsSeriesInRows = false;
                //Set a Chart Title.
                chart.ChartTitle = "Doughnut Chart";
                //Set Datalabels.
                IOfficeChartSerie series1 = chart.Series[0];
                IOfficeChartSerie series2 = chart.Series[1];

                series1.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
                series2.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
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
