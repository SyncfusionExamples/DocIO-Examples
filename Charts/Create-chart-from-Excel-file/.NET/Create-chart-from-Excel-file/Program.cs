using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;
using System.IO;

namespace Create_chart_from_Excel_file
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
                //Loads the excel file as stream.
                Stream excelStream = File.OpenRead(Path.GetFullPath(@"Data/Excel_Template.xlsx"));
                //Creates and Appends chart to the paragraph with excel stream as parameter.
                WChart chart = paragraph.AppendChart(excelStream, 1, "B2:C6", 470, 300);
                //Sets chart type and title.
                chart.ChartType = OfficeChartType.Column_Clustered;
                chart.ChartTitle = "Purchase Details";
                chart.ChartTitleArea.FontName = "Calibri";
                chart.ChartTitleArea.Size = 14;
                chart.ChartArea.Border.LinePattern = OfficeChartLinePattern.None;
                //Sets name to chart series.         
                chart.Series[0].Name = "Sum of Purchases";
                chart.Series[1].Name = "Sum of Future Expenses";
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
