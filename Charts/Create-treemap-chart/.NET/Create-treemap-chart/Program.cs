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

                // Create and append TreeMap chart to the paragraph.
                WChart chart = paragraph.AppendChart(446, 270);
                chart.ChartType = OfficeChartType.TreeMap;

                // Set chart title.
                chart.ChartTitle = "Food Sales - Treemap Chart";

                // Set headers.
                chart.ChartData.SetValue(1, 1, "Meal");
                chart.ChartData.SetValue(1, 2, "Category");
                chart.ChartData.SetValue(1, 3, "Item");
                chart.ChartData.SetValue(1, 4, "Sales");

                // Add data rows.
                string[,] data = {
                {"Breakfast", "Beverage", "coffee", "20"},
                {"Breakfast", "Beverage", "tea", "9"},
                {"Breakfast", "Food", "waffles", "12"},
                {"Breakfast", "Food", "pancakes", "35"},
                {"Breakfast", "Food", "eggs", "24"},
                {"Lunch", "Beverage", "coffee", "10"},
                {"Lunch", "Beverage", "iced tea", "45"},
                {"Lunch", "Food", "soup", "16"},
                {"Lunch", "Food", "sandwich", "36"},
                {"Lunch", "Food", "salad", "70"},
                {"Lunch", "Food", "pie", "45"},
                {"Lunch", "Food", "cookie", "25"}
            };

                for (int i = 0; i < data.GetLength(0); i++)
                {
                    chart.ChartData.SetValue(i + 2, 1, data[i, 0]);
                    chart.ChartData.SetValue(i + 2, 2, data[i, 1]);
                    chart.ChartData.SetValue(i + 2, 3, data[i, 2]);
                    chart.ChartData.SetValue(i + 2, 4, int.Parse(data[i, 3]));
                }

                chart.DataRange = chart.ChartData[2, 1, 13, 4];
                // Set DataLabels.
                IOfficeChartSerie serie = chart.Series[0];
                serie.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;

                // Set legend.
                chart.HasLegend = true;
                chart.Legend.Position = OfficeLegendPosition.Bottom;

                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Save the Word document to the file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
