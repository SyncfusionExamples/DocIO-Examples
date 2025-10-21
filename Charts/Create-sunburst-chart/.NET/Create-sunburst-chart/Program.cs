using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;

namespace Create_pie_chart
{
    class Program
    {
        static void Main(string[] args)
        {
            //Create a new Word document.
            using (WordDocument document = new WordDocument())
            {
                // Add a section & a paragraph to the document.
                IWSection section = document.AddSection();
                IWParagraph paragraph = section.AddParagraph();

                // Create and append Sunburst chart to the paragraph.
                WChart chart = paragraph.AppendChart(446, 270);
                chart.ChartType = OfficeChartType.SunBurst;

                // Set chart title.
                chart.ChartTitle = "Sales by Annual - Sunburst Chart";

                // Set headers.
                chart.ChartData.SetValue(1, 1, "Quarter");
                chart.ChartData.SetValue(1, 2, "Month");
                chart.ChartData.SetValue(1, 3, "Week");
                chart.ChartData.SetValue(1, 4, "Sales");

                // Add data rows.
                string[,] data = {
                {"1st", "Jan", "", "3.5"},
                {"1st", "Feb", "Week 1", "1.2"},
                {"1st", "Feb", "Week 2", "0.8"},
                {"1st", "Feb", "Week 3", "0.6"},
                {"1st", "Feb", "Week 4", "0.5"},
                {"1st", "Mar", "", "1.7"},
                {"2nd", "Apr", "", "1.1"},
                {"2nd", "May", "", "0.8"},
                {"2nd", "Jun", "", "0.8"},
                {"3rd", "Jul", "", "1"},
                {"3rd", "Aug", "", "0.7"},
                {"3rd", "Sep", "", "0.9"},
                {"4th", "Oct", "", "2"},
                {"4th", "Nov", "", "2"},
                {"4th", "Dec", "", "2"}
            };

                for (int i = 0; i < data.GetLength(0); i++)
                {
                    chart.ChartData.SetValue(i + 2, 1, data[i, 0]);
                    chart.ChartData.SetValue(i + 2, 2, data[i, 1]);
                    chart.ChartData.SetValue(i + 2, 3, data[i, 2]);
                    chart.ChartData.SetValue(i + 2, 4, float.Parse(data[i, 3]));
                }

                // Set data range and hierarchy.
                chart.DataRange = chart.ChartData[2, 1, data.GetLength(0) + 1, 4];
                // Set DataLabels.
                IOfficeChartSerie serie = chart.Series[0];
                serie.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;

                // Set legend.
                chart.HasLegend = true;
                chart.Legend.Position = OfficeLegendPosition.Bottom;

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
