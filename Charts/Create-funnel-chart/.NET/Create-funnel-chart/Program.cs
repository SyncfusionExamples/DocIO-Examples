using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;

namespace Create_funnel_chart
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
                // Add a paragraph to the section.
                IWParagraph paragraph = section.AddParagraph();

                // Create and append the chart to the paragraph.
                WChart chart = paragraph.AppendChart(446, 270);
                // Set chart type.
                chart.ChartType = OfficeChartType.Funnel;

                // Set chart title.
                chart.ChartTitle = "Sales Stages - Funnel Chart";

                // Set headers.
                chart.ChartData.SetValue(1, 1, "Stage");
                chart.ChartData.SetValue(1, 2, "Amount");

                // Add data rows.
                string[,] data = {
        {"Prospects", "500"},
        {"Qualified prospects", "425"},
        {"Needs analysis", "200"},
        {"Price quotes", "150"},
        {"Negotiations", "100"},
        {"Closed sales", "90"}
    };

                for (int i = 0; i < data.GetLength(0); i++)
                {
                    chart.ChartData.SetValue(i + 2, 1, data[i, 0]);
                    chart.ChartData.SetValue(i + 2, 2, int.Parse(data[i, 1]));
                }

                // Set data range.
                chart.DataRange = chart.ChartData[2, 1, 7, 2];

                IOfficeChartSerie series = chart.Series[0];
                series.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
                series.DataPoints.DefaultDataPoint.DataLabels.Position = OfficeDataLabelPosition.Inside;

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
