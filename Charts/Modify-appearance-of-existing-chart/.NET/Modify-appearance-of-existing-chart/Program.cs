using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;
using System.IO;

namespace Modify_appearance_of_existing_chart
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                //Opens an existing Word document.
                using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
                {
                    //Gets the paragraph.
                    WParagraph paragraph = document.LastParagraph;
                    //Gets the chart entity.
                    WChart chart = paragraph.ChildEntities[0] as WChart;
                    //Modifies the chart height and width.
                    chart.Height = 300;
                    chart.Width = 500;
                    //Changes the title.
                    chart.ChartTitle = "New title";
                    //Changes the series name of first chart series.
                    chart.Series[0].Name = "Modified series name";
                    //Hides the category labels.
                    chart.CategoryLabelLevel = OfficeCategoriesLabelLevel.CategoriesLabelLevelNone;
                    //Shows data Table.
                    chart.HasDataTable = true;
                    //Formats the chart area.
                    IOfficeChartFrameFormat chartArea = chart.ChartArea;
                    //Sets border line pattern, color, line weight.
                    chartArea.Border.LinePattern = OfficeChartLinePattern.Solid;
                    chartArea.Border.LineColor = Syncfusion.Drawing.Color.Blue;
                    chartArea.Border.LineWeight = OfficeChartLineWeight.Hairline;
                    //Sets fill type and fill colors.
                    chartArea.Fill.FillType = OfficeFillType.Gradient;
                    chartArea.Fill.GradientColorType = OfficeGradientColor.TwoColor;
                    chartArea.Fill.BackColor = Syncfusion.Drawing.Color.FromArgb(205, 217, 234);
                    chartArea.Fill.ForeColor = Syncfusion.Drawing.Color.White;
                    //Plots Area.
                    IOfficeChartFrameFormat chartPlotArea = chart.PlotArea;
                    //Plots area border settings - line pattern, color, weight.
                    chartPlotArea.Border.LinePattern = OfficeChartLinePattern.Solid;
                    chartPlotArea.Border.LineColor = Syncfusion.Drawing.Color.Blue;
                    chartPlotArea.Border.LineWeight = OfficeChartLineWeight.Hairline;
                    //Sets fill type and color.
                    chartPlotArea.Fill.FillType = OfficeFillType.Gradient;
                    chartPlotArea.Fill.GradientColorType = OfficeGradientColor.TwoColor;
                    chartPlotArea.Fill.BackColor = Syncfusion.Drawing.Color.FromArgb(205, 217, 234);
                    chartPlotArea.Fill.ForeColor = Syncfusion.Drawing.Color.White;
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
}
