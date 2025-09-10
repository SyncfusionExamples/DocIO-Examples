using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;

namespace Create_box_and_whisker_chart
{
    class Program
    {
        static void Main(string[] args)
        {
            using (WordDocument document = new WordDocument())
            {
                IWSection section = document.AddSection();
                IWParagraph paragraph = section.AddParagraph();

                WChart chart = paragraph.AppendChart(500f, 300f);
                chart.ChartType = OfficeChartType.BoxAndWhisker;
                chart.ChartTitle = "Box and whisker chart";

                // Set headers
                chart.ChartData.SetValue(1, 1, "Course");
                chart.ChartData.SetValue(1, 2, "School A");
                chart.ChartData.SetValue(1, 3, "School B");
                chart.ChartData.SetValue(1, 4, "School C");

                // Add data rows
                string[,] data = {
        {"English", "63", "53", "45"},
        {"Physics", "61", "55", "65"},
        {"English", "63", "50", "65"},
        {"Math", "62", "51", "64"},
        {"English", "46", "53", "66"},
        {"English", "58", "56", "67"},
        {"Math", "60", "51", "67"},
        {"Math", "62", "53", "66"},
        {"English", "63", "54", "64"},
        {"English", "63", "52", "67"},
        {"Physics", "60", "56", "64"},
        {"English", "60", "56", "67"},
        {"Math", "61", "56", "45"},
        {"Math", "63", "58", "64"},
        {"English", "59", "54", "65"}
    };

                for (int i = 0; i < data.GetLength(0); i++)
                {
                    chart.ChartData.SetValue(i + 2, 1, data[i, 0]); // Course name
                    chart.ChartData.SetValue(i + 2, 2, int.Parse(data[i, 1])); // School A
                    chart.ChartData.SetValue(i + 2, 3, int.Parse(data[i, 2])); // School B
                    chart.ChartData.SetValue(i + 2, 4, int.Parse(data[i, 3])); // School C
                }

                // Set data range and chart properties
                chart.DataRange = chart.ChartData[2, 1, data.GetLength(0) + 1, 4];

                chart.PrimaryCategoryAxis.Title = "Subjects";
                chart.PrimaryValueAxis.Title = "Scores";
                chart.PrimaryValueAxis.MinimumValue = 0;
                chart.PrimaryValueAxis.MaximumValue = 70;

                IOfficeChartSerie series = chart.Series[0];
                series.SerieFormat.ShowOutlierPoints = true;
                series.SerieFormat.ShowMeanMarkers = true;

                series = chart.Series[1];
                series.SerieFormat.ShowOutlierPoints = true;
                series.SerieFormat.ShowMeanMarkers = true;

                series = chart.Series[2];
                series.SerieFormat.ShowOutlierPoints = true;
                series.SerieFormat.ShowMeanMarkers = true;

                // Set legend.
                chart.HasLegend = true;
                chart.Legend.Position = OfficeLegendPosition.Bottom;

                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
