using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using Syncfusion.OfficeChart;
using System.IO;

namespace Remove_chart_title
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
                chart.ChartTitle = string.Empty;

                //Sets data for chart.
                chart.ChartData.SetValue(1, 1, "");
                chart.ChartData.SetValue(1, 2, "Sales");
                chart.ChartData.SetValue(2, 1, "Phyllis Lapin");
                chart.ChartData.SetValue(2, 2, 141.396);
                chart.ChartData.SetValue(3, 1, "Stanley Hudson");
                chart.ChartData.SetValue(3, 2, 80.368);
                //Creates a new chart series with the name “Sales”.
                IOfficeChartSerie pieSeries = chart.Series.Add("Sales");
                pieSeries.Values = chart.ChartData[2, 2, 3, 2];
                //Sets category labels.
                chart.PrimaryCategoryAxis.CategoryLabels = chart.ChartData[2, 1, 3, 1];
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
