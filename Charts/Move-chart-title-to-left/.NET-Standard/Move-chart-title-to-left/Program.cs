using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;
using System.IO;

namespace Move_chart_title_to_left
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
                chart.ChartType = OfficeChartType.Pie;

                //Set position for title area.
                chart.ChartTitleArea.Layout.ManualLayout.LeftMode = LayoutModes.edge;
                chart.ChartTitleArea.Layout.ManualLayout.TopMode = LayoutModes.edge;
                chart.ChartTitleArea.Layout.ManualLayout.Left = 0.041214980185031891;
                chart.ChartTitleArea.Layout.ManualLayout.Top = 0.0560000017285347;

                //Set chart data.
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
                //Set chart title.
                chart.ChartTitle = "Sales Report";
                //Create a new chart series with the name “Sales”.
                IOfficeChartSerie pieSeries = chart.Series.Add("Sales");
                pieSeries.Values = chart.ChartData[2, 2, 11, 2];
                //Set data label.
                pieSeries.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
                pieSeries.DataPoints.DefaultDataPoint.DataLabels.Position = OfficeDataLabelPosition.Outside;
                //Set category labels.
                chart.PrimaryCategoryAxis.CategoryLabels = chart.ChartData[2, 1, 11, 1];
                //Set legend.
                chart.Legend.Position = OfficeLegendPosition.Left;
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
