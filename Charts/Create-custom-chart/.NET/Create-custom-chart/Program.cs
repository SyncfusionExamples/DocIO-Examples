using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;
using System.IO;

namespace Create_custom_chart
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
                //Inputs data for chart.
                object[][] data = new object[6][];
                for (int i = 0; i < 6; i++)
                    data[i] = new object[3];
                data[0][0] = "";
                data[1][0] = "Camembert Pierrot";
                data[2][0] = "Alice Mutton";
                data[3][0] = "Roasted Tigers";
                data[4][0] = "Orange Shake";
                data[5][0] = "Dried Apples";
                data[0][1] = "Sum of Purchases";
                data[1][1] = 286;
                data[2][1] = 680;
                data[3][1] = 288;
                data[4][1] = 200;
                data[5][1] = 731;
                data[0][2] = "Sum of Future Expenses";
                data[1][2] = 1300;
                data[2][2] = 700;
                data[3][2] = 1280;
                data[4][2] = 1200;
                data[5][2] = 2660;
                //Creates and Appends chart to the paragraph.
                WChart chart = paragraph.AppendChart(data, 470, 300);
                //Sets chart type and title.
                chart.ChartTitle = "Purchase Details";
                chart.ChartTitleArea.FontName = "Calibri";
                chart.ChartTitleArea.Size = 14;
                chart.ChartArea.Border.LinePattern = OfficeChartLinePattern.None;
                //Sets series type.
                chart.Series[0].SerieType = OfficeChartType.Line_Markers;
                chart.Series[1].SerieType = OfficeChartType.Bar_Clustered;
                chart.PrimaryCategoryAxis.Title = "Products";
                chart.PrimaryValueAxis.Title = "In Dollars";
                //Sets position of legend.
                chart.Legend.Position = OfficeLegendPosition.Bottom;
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
