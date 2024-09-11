using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using Syncfusion.OfficeChart;
using System.IO;

namespace Create_chart_from_scratch
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new instance of WordDocument (Empty Word Document).
            using (WordDocument document = new WordDocument())
            {
                //Adds section to the document.
                IWSection sec = document.AddSection();
                //Adds paragraph to the section.
                IWParagraph paragraph = sec.AddParagraph();
                //Creates and Appends chart to the paragraph.
                WChart chart = paragraph.AppendChart(446, 270);
                //Sets chart type.
                chart.ChartType = OfficeChartType.Pie;
                //Sets chart title.
                chart.ChartTitle = "Best Selling Products";
                chart.ChartTitleArea.FontName = "Calibri";
                chart.ChartTitleArea.Size = 14;
                //Sets data for chart.
                chart.ChartData.SetValue(1, 1, "");
                chart.ChartData.SetValue(1, 2, "Sales");
                chart.ChartData.SetValue(2, 1, "Phyllis Lapin");
                chart.ChartData.SetValue(2, 2, 141.396);
                chart.ChartData.SetValue(3, 1, "Stanley Hudson");
                chart.ChartData.SetValue(3, 2, 80.368);
                chart.ChartData.SetValue(4, 1, "Bernard Shah");
                chart.ChartData.SetValue(4, 2, 71.155);
                chart.ChartData.SetValue(5, 1, "Patricia Lincoln");
                chart.ChartData.SetValue(5, 2, 47.234);
                chart.ChartData.SetValue(6, 1, "Camembert Pierrot");
                chart.ChartData.SetValue(6, 2, 46.825);
                chart.ChartData.SetValue(7, 1, "Thomas Hardy");
                chart.ChartData.SetValue(7, 2, 42.593);
                chart.ChartData.SetValue(8, 1, "Hanna Moos");
                chart.ChartData.SetValue(8, 2, 41.819);
                chart.ChartData.SetValue(9, 1, "Alice Mutton");
                chart.ChartData.SetValue(9, 2, 32.698);
                chart.ChartData.SetValue(10, 1, "Christina Berglund");
                chart.ChartData.SetValue(10, 2, 29.171);
                chart.ChartData.SetValue(11, 1, "Elizabeth Lincoln");
                chart.ChartData.SetValue(11, 2, 25.696);
                //Creates a new chart series with the name “Sales”.
                IOfficeChartSerie pieSeries = chart.Series.Add("Sales");
                pieSeries.Values = chart.ChartData[2, 2, 11, 2];
                //Sets data label.
                pieSeries.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
                pieSeries.DataPoints.DefaultDataPoint.DataLabels.Position = OfficeDataLabelPosition.Outside;
                //Sets background color.
                chart.ChartArea.Fill.ForeColor = Syncfusion.Drawing.Color.FromArgb(242, 242, 242);
                chart.PlotArea.Fill.ForeColor = Syncfusion.Drawing.Color.FromArgb(242, 242, 242);
                chart.ChartArea.Border.LinePattern = OfficeChartLinePattern.None;
                //Sets category labels.
                chart.PrimaryCategoryAxis.CategoryLabels = chart.ChartData[2, 1, 11, 1];
                //Creates file stream.
                using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite))
                {
                    //Saves the Word document to file stream.
                    document.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}
