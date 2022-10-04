using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;
using System.IO;

namespace Create_bar_clustered_chart
{
    class Program
    {
        static void Main(string[] args)
        {
            //Create a new Word document.
            using (WordDocument document = new WordDocument())
            {
                // Add a section to the document.
                IWSection section = document.AddSection();
                //Add a paragraph to the section.
                IWParagraph paragraph = section.AddParagraph();
                //Create and append the chart to the paragraph.
                WChart chart = paragraph.AppendChart(446, 270);
                //Set chart type.
                chart.ChartType = OfficeChartType.Bar_Clustered;
                //Assign data.
                AddChartData(chart);
                //Set a Chart Title.
                chart.ChartTitle = "Bar clustered chart";
                //Set Datalabels.
                IOfficeChartSerie series1 = chart.Series.Add("Series 1");
                //Set the data range of chart series – start row, start column, end row and end column.
                series1.Values = chart.ChartData[2, 2, 4, 2];
                IOfficeChartSerie series2 = chart.Series.Add("Series 2");
                //Set the data range of chart series start row, start column, end row and end column.
                series2.Values = chart.ChartData[2, 3, 4, 3];
                IOfficeChartSerie series3 = chart.Series.Add("Series 3");
                //Set the data range of chart series start row, start column, end row and end column.
                series3.Values = chart.ChartData[2, 4, 4, 4];
                //Set the data range of the category axis.
                chart.PrimaryCategoryAxis.CategoryLabels = chart.ChartData[2, 1, 4, 1];

                //Set legend position.
                chart.Legend.Position = OfficeLegendPosition.Bottom;
                //Save the file in the given path.
                Stream docStream = File.Create(Path.GetFullPath(@"../../../Sample.docx"));
                document.Save(docStream, FormatType.Docx);
                docStream.Dispose();
            }
        }
        /// <summary>
        /// Set the values for the chart.
        /// </summary>
        private static void AddChartData(WChart chart)
        {
            //Set the value for chart data.
            chart.ChartData.SetValue(1, 2, "Series1");
            chart.ChartData.SetValue(1, 3, "Series2");
            chart.ChartData.SetValue(1, 4, "Series3");

            chart.ChartData.SetValue(2, 1, "Category1");
            chart.ChartData.SetValue(2, 2, 5);
            chart.ChartData.SetValue(2, 3, 4);
            chart.ChartData.SetValue(2, 4, 3);

            chart.ChartData.SetValue(3, 1, "Category2");
            chart.ChartData.SetValue(3, 2, 4);
            chart.ChartData.SetValue(3, 3, 5);
            chart.ChartData.SetValue(3, 4, 2);

            chart.ChartData.SetValue(4, 1, "Category3");
            chart.ChartData.SetValue(4, 2, 4);
            chart.ChartData.SetValue(4, 3, 4);
            chart.ChartData.SetValue(4, 4, 3);
        }
    }   
}
