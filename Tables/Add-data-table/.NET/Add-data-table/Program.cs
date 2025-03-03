using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;
using System.IO;

namespace AddDataTable
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new Word document.
            WordDocument document = new WordDocument();

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

            // Create a file stream to save the document.
            using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite))
            {
                // Save the Word document to the file stream.
                document.Save(outputFileStream, FormatType.Docx);
            }
        }
    }
}
