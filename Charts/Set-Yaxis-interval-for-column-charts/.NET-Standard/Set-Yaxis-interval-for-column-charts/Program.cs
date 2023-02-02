using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;
using System.Collections.Generic;
using System.IO;

namespace Set_Yaxis_interval_for_column_charts
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
                chart.ChartType = OfficeChartType.Column_Clustered;
                //Set chart data.
                chart.ChartData.SetValue(1, 1, "Items");
                chart.ChartData.SetValue(1, 2, "Amount(in $)");
                chart.ChartData.SetValue(1, 3, "Count");

                chart.ChartData.SetValue(2, 1, "Beverages");
                chart.ChartData.SetValue(2, 2, 277);
                chart.ChartData.SetValue(2, 3, 925);

                chart.ChartData.SetValue(3, 1, "Condiments");
                chart.ChartData.SetValue(3, 2, 177);
                chart.ChartData.SetValue(3, 3, 378);

                chart.ChartData.SetValue(4, 1, "Confections");
                chart.ChartData.SetValue(4, 2, 387);
                chart.ChartData.SetValue(4, 3, 880);

                chart.ChartData.SetValue(5, 1, "Dairy Products");
                chart.ChartData.SetValue(5, 2, 1008);
                chart.ChartData.SetValue(5, 3, 581);

                chart.ChartData.SetValue(6, 1, "Grains/Cereals");
                chart.ChartData.SetValue(6, 2, 1500);
                chart.ChartData.SetValue(6, 3, 189);
                //Set chart series in the column for assigned data region.
                chart.IsSeriesInRows = false;

                IOfficeChartSerie serie1 = chart.Series.Add("Amount(in $)");
                //Set the data range of chart series – start row, start column, end row, end column.
                serie1.Values = chart.ChartData[2, 2, 6, 2];
                IOfficeChartSerie serie2 = chart.Series.Add("Count");
                //Set the data range of chart series – start row, start column, end row, end column.
                serie2.Values = chart.ChartData[2, 3, 6, 3];
                //Set Datalabels.
                chart.PrimaryCategoryAxis.CategoryLabels = chart.ChartData[2, 1, 6, 1];
                //Apply chart elements.
                //Set Chart Title.
                chart.ChartTitle = "Clustered Column Chart";

                //Set the minimum and maximum value of the Y-Axis.
                chart.PrimaryValueAxis.MinimumValue = 0;
                chart.PrimaryValueAxis.MaximumValue = 1600;

                //Sets the interval for Y-Axis
                chart.PrimaryValueAxis.MajorUnit = 200;

                serie1.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
                serie2.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
                serie1.DataPoints.DefaultDataPoint.DataLabels.Position = OfficeDataLabelPosition.Center;
                serie2.DataPoints.DefaultDataPoint.DataLabels.Position = OfficeDataLabelPosition.Center;
                //Set Legend.
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
