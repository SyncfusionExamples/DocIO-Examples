using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using Syncfusion.OfficeChart;
using System.IO;

namespace Add_data_table
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates a new instance of WordDocument.
            using (WordDocument document = new WordDocument())
            {
                // Add a section to the document.
                IWSection section = document.AddSection();
                // Add a paragraph to the section.
                IWParagraph paragraph = section.AddParagraph();
                // Create and append the chart to the paragraph.
                WChart chart = paragraph.AppendChart(446, 270);
                // Set chart type.
                chart.ChartType = OfficeChartType.Column_Clustered;
                // Set chart data.
                chart.ChartData.SetValue(2, 1, "Apples");
                chart.ChartData.SetValue(3, 1, "Grapes");
                chart.ChartData.SetValue(4, 1, "Banana");
                chart.ChartData.SetValue(1, 2, "Joey");
                chart.ChartData.SetValue(2, 2, 5);
                chart.ChartData.SetValue(3, 2, 4);
                chart.ChartData.SetValue(4, 2, 4);
                // Define the data range for the chart.
                chart.DataRange = chart.ChartData[1, 1, 4, 2];
                // Add data table to the chart.
                chart.HasDataTable = true;
                IOfficeChartDataTable officeChartDataTable = chart.DataTable;
                // Customize the data table appearance.
                officeChartDataTable.ShowSeriesKeys = true;
                officeChartDataTable.HasBorders = true;
                officeChartDataTable.HasHorzBorder = true;
                officeChartDataTable.HasVertBorder = true;
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
