using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;

namespace Create_histogram_chart
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new Word document.
            using (WordDocument document = new WordDocument())
            {
                IWSection sec = document.AddSection();
                IWParagraph paragraph = sec.AddParagraph();

                // Create and append chart
                WChart chart = paragraph.AppendChart(446, 270);
                chart.ChartType = OfficeChartType.Histogram;

                // Set chart title
                chart.ChartTitle = "Test Scores (in Histogram)";
                chart.ChartTitleArea.FontName = "Calibri";
                chart.ChartTitleArea.Size = 14;

                // Set chart data
                chart.ChartData.SetValue(1, 1, "Test Score");
                int[] scores = { 20, 35, 40, 55, 80, 60, 61, 85, 80, 64, 80, 75 };
                for (int i = 0; i < scores.Length; i++)
                {
                    chart.ChartData.SetValue(i + 2, 1, scores[i]);
                }
                // Set region of chart data.
                chart.DataRange = chart.ChartData[2, 1, 13, 1];

                // Set DataLabels.
                IOfficeChartSerie series = chart.Series[0];
                series.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
                series.Name = "Score";
                series.SerieFormat.CommonSerieOptions.GapWidth = 3;

                // Set legend.
                chart.HasLegend = true;
                chart.Legend.Position = OfficeLegendPosition.Bottom;

                // Create a file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    // Save the Word document to the file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
