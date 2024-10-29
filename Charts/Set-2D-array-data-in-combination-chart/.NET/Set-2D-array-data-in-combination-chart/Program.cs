using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;

namespace Set_2D_array_data_in_combination_chart
{
    class Program
    {
        static void Main(string[] args)
        {
            //Define a 2D array with dimensions [3, 13]
            string[,] data = new string[3, 13]
            {
                { "Month", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" },
                { "Rainy Days", "12", "11", "10", "9", "8", "6", "4", "6", "7", "8", "10", "11" },
                { "Profit", "3574", "4708", "5332", "6693", "8843", "12347", "15180", "11198", "9739", "9846", "6620", "5085" }
            };
            //Create a new Word document.
            using (WordDocument document = new WordDocument())
            {
                //Add a section to the document.
                IWSection section = document.AddSection();
                //Add a paragraph to the section.
                IWParagraph paragraph = section.AddParagraph();
                //Create and append the chart to the paragraph.
                WChart chart = paragraph.AppendChart(446, 270);
                //Set chart data using values from the 2D array.
                for (int i = 0; i < data.GetLength(0); i++)
                {
                    for (int j = 0; j < data.GetLength(1); j++)
                    {
                        string value = data[i, j];
                        //Try to parse the value as an integer, if possible
                        if (int.TryParse(value, out int intValue))
                        {
                            //Set the integer value.
                            chart.ChartData.SetValue(j + 1, i + 1, intValue);
                        }
                        else
                        {
                            //Set the string value.
                            chart.ChartData.SetValue(j + 1, i + 1, value);
                        }
                    }
                }
                //Set region of Chart data.
                chart.DataRange = chart.ChartData[1, 1, 13, 3];
                //Set chart series in the column for assigned data region.
                chart.IsSeriesInRows = false;
                //Set a Chart Title.
                chart.ChartTitle = "Combination Chart";
                //Set Datalabels.
                IOfficeChartSerie series1 = chart.Series[0];
                IOfficeChartSerie series2 = chart.Series[1];
                //Set Serie type.
                series1.SerieType = OfficeChartType.Column_Clustered;
                series2.SerieType = OfficeChartType.Line;
                series2.UsePrimaryAxis = false;
                //Set Datalabels.
                series1.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
                //Set legend.
                chart.HasLegend = true;
                chart.Legend.Position = OfficeLegendPosition.Bottom;
                //Set chart type.
                chart.ChartType = OfficeChartType.Combination_Chart;
                //Set secondary axis on right side.
                chart.SecondaryValueAxis.TickLabelPosition = OfficeTickLabelPosition.TickLabelPosition_High;
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
