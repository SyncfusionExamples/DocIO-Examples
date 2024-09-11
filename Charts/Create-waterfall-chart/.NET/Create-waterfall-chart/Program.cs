using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;

namespace Create_waterfall_chart
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Adds section to the document.
                IWSection sec = document.AddSection();
                //Adds paragraph to the section.
                IWParagraph paragraph = sec.AddParagraph();
                //Creates and Appends chart to the paragraph.
                WChart chart = paragraph.AppendChart(446, 270);
                //Sets chart type.
                chart.ChartType = OfficeChartType.WaterFall;
                //Sets data range.
                chart.DataRange = chart.ChartData[1, 1, 8, 2];
                chart.IsSeriesInRows = false;
                chart.ChartData.SetValue(2, 1, "Start");
                chart.ChartData.SetValue(2, 2, 120000);
                chart.ChartData.SetValue(3, 1, "Product Revenue");
                chart.ChartData.SetValue(3, 2, 570000);
                chart.ChartData.SetValue(4, 1, "Service Revenue");
                chart.ChartData.SetValue(4, 2, 230000);
                chart.ChartData.SetValue(5, 1, "Positive Balance");
                chart.ChartData.SetValue(5, 2, 920000);
                chart.ChartData.SetValue(6, 1, "Fixed Costs");
                chart.ChartData.SetValue(6, 2, -345000);
                chart.ChartData.SetValue(7, 1, "Variable Costs");
                chart.ChartData.SetValue(7, 2, -230000);
                chart.ChartData.SetValue(8, 1, "Total");
                chart.ChartData.SetValue(8, 2, 345000);
                //Data point settings as total in chart.
                IOfficeChartSerie series = chart.Series[0];
                chart.Series[0].DataPoints[3].SetAsTotal = true;
                chart.Series[0].DataPoints[6].SetAsTotal = true;
                //Showing the connector lines between data points.
                chart.Series[0].SerieFormat.ShowConnectorLines = true;
                //Set the chart title.
                chart.ChartTitle = "Company Profit (in USD)";
                //Formatting data label and legend option.
                chart.Series[0].DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
                chart.Series[0].DataPoints.DefaultDataPoint.DataLabels.Size = 8;
                chart.Legend.Position = OfficeLegendPosition.Right;
                chart.ChartArea.Border.LinePattern = OfficeChartLinePattern.None;
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
