using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;
using System.IO;

namespace Create_stacked_bar_chart
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
                chart.ChartType = OfficeChartType.Bar_Stacked;
                //Assign data.
                AddChartData(chart);
                //Set chart series in the column for assigned data region
                chart.IsSeriesInRows = false;
                //Set a Chart Title.
                chart.ChartTitle = "Stacked Bar Chart";
                //Set Datalabels.
                IOfficeChartSerie series1 = chart.Series.Add("Series 1");
                //Set the data range of chart series – start row, start column, end row, and end column.
                series1.Values = chart.ChartData[2, 2, 4, 2];
                IOfficeChartSerie series2 = chart.Series.Add("Series 2");
                //Set the data range of chart series start row, start column, end row, and end column.
                series2.Values = chart.ChartData[2, 3, 4, 3];
                IOfficeChartSerie series3 = chart.Series.Add("Series 3");
                //Set the data range of chart series start row, start column, end row, and end column.
                series3.Values = chart.ChartData[2, 4, 4, 4];
                //Set the data range of the category axis.
                chart.PrimaryCategoryAxis.CategoryLabels = chart.ChartData[2, 1, 4, 1];

                series1.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
                series2.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
                series3.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
                series1.DataPoints.DefaultDataPoint.DataLabels.Position = OfficeDataLabelPosition.Center;
                series2.DataPoints.DefaultDataPoint.DataLabels.Position = OfficeDataLabelPosition.Center;
                series3.DataPoints.DefaultDataPoint.DataLabels.Position = OfficeDataLabelPosition.Center;

                //Set legend.
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
