using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;

namespace Create_pareto_chart
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new Word document.
            using (WordDocument document = new WordDocument())
            {
                // Add a section and a paragraph.
                IWSection section = document.AddSection();
                IWParagraph paragraph = section.AddParagraph();
                // Create and append a Pareto chart.
                WChart chart = paragraph.AppendChart(446, 270);
                chart.ChartType = OfficeChartType.Pareto;
                chart.ChartTitle = "Monthly Expenses - Pareto Chart";
                chart.ChartTitleArea.FontName = "Calibri";
                chart.ChartTitleArea.Size = 14;

                // Populate chart data in the internal worksheet.
                chart.ChartData.SetValue(1, 1, "Expenses");
                chart.ChartData.SetValue(1, 2, "Amount");

                chart.ChartData.SetValue(2, 1, "Rent");
                chart.ChartData.SetValue(2, 2, 2300);

                chart.ChartData.SetValue(3, 1, "Car payment");
                chart.ChartData.SetValue(3, 2, 1200);

                chart.ChartData.SetValue(4, 1, "Groceries");
                chart.ChartData.SetValue(4, 2, 900);

                chart.ChartData.SetValue(5, 1, "Electricity");
                chart.ChartData.SetValue(5, 2, 600);

                chart.ChartData.SetValue(6, 1, "Gas");
                chart.ChartData.SetValue(6, 2, 500);

                chart.ChartData.SetValue(7, 1, "House loan");
                chart.ChartData.SetValue(7, 2, 300);

                chart.ChartData.SetValue(8, 1, "Wifi bill");
                chart.ChartData.SetValue(8, 2, 200);

                // Define the data range for the chart.
                chart.DataRange = chart.ChartData[1, 1, 8, 2];

                // Format the right Y-axis to show percentage.
                IOfficeChartAxis rightAxis = chart.SecondaryValueAxis;
                rightAxis.Title = "Cumulative Percentage";

                // Set gap width for the first series (bars)
                IOfficeChartSerie series = chart.Series[0];
                series.SerieFormat.CommonSerieOptions.GapWidth = 3;
                series.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;

                // Set legend.
                chart.HasLegend = true;
                chart.Legend.Position = OfficeLegendPosition.Bottom;

                // Save the document.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
