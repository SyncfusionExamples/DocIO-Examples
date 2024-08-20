using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;
using System.IO;

namespace Create_100_stacked_bar_cone_chart
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
                //Set chart data.
                chart.ChartData.SetValue(1, 1, "Fruits");
                chart.ChartData.SetValue(2, 1, "Apples");
                chart.ChartData.SetValue(3, 1, "Grapes");
                chart.ChartData.SetValue(4, 1, "Bananas");
                chart.ChartData.SetValue(5, 1, "Oranges");
                chart.ChartData.SetValue(6, 1, "Melons");
                chart.ChartData.SetValue(1, 2, "Joey");
                chart.ChartData.SetValue(2, 2, 5);
                chart.ChartData.SetValue(3, 2, 4);
                chart.ChartData.SetValue(4, 2, 4);
                chart.ChartData.SetValue(5, 2, 2);
                chart.ChartData.SetValue(6, 2, 2);
                chart.ChartData.SetValue(1, 3, "Matthew");
                chart.ChartData.SetValue(2, 3, 3);
                chart.ChartData.SetValue(3, 3, 5);
                chart.ChartData.SetValue(4, 3, 4);
                chart.ChartData.SetValue(5, 3, 1);
                chart.ChartData.SetValue(6, 3, 7);
                chart.ChartData.SetValue(1, 4, "Peter");
                chart.ChartData.SetValue(2, 4, 2);
                chart.ChartData.SetValue(3, 4, 2);
                chart.ChartData.SetValue(4, 4, 3);
                chart.ChartData.SetValue(5, 4, 5);
                chart.ChartData.SetValue(6, 4, 6);
                //Set region of Chart data.
                chart.DataRange = chart.ChartData[1, 1, 6, 4];
                //Set chart series in the column for assigned data region
                chart.IsSeriesInRows = false;
                //Set a Chart Title.
                chart.ChartTitle = "100% Stacked Bar Cone Chart";
                //Set Datalabels.
                IOfficeChartSerie series1 = chart.Series[0];
                IOfficeChartSerie series2 = chart.Series[1];
                IOfficeChartSerie series3 = chart.Series[2];

                series1.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
                series2.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
                series3.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
                //Set legend.
                chart.HasLegend = true;
                chart.Legend.Position = OfficeLegendPosition.Bottom;
                //Set chart type.
                chart.ChartType = OfficeChartType.Cone_Bar_Stacked_100;
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
