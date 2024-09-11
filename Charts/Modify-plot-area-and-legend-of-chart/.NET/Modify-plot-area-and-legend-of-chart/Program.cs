using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;
using System.IO;

namespace Modify_plot_area_and_legend_of_chart
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
                    //Sets border settings - line color, pattern, weight, transparency.
                    chart.PlotArea.Border.AutoFormat = false;
                    chart.PlotArea.Border.IsAutoLineColor = false;
                    chart.PlotArea.Border.LineColor = Syncfusion.Drawing.Color.Blue;
                    chart.PlotArea.Border.LinePattern = OfficeChartLinePattern.DashDot;
                    chart.PlotArea.Border.LineWeight = OfficeChartLineWeight.Wide;
                    chart.PlotArea.Border.Transparency = 0.6;
                    //Sets the plot area’s fill type, color.
                    chart.PlotArea.Fill.FillType = OfficeFillType.SolidColor;
                    chart.PlotArea.Fill.ForeColor = Syncfusion.Drawing.Color.LightPink;
                    //Sets the plot area shadow presence.
                    chart.PlotArea.Shadow.ShadowInnerPresets = Office2007ChartPresetsInner.InsideDiagonalTopLeft;
                    //Sets the legend position.
                    chart.Legend.Position = OfficeLegendPosition.Left;
                    //Sets the layout inclusion.
                    chart.Legend.IncludeInLayout = true;
                    //Sets the legend border format - color, pattern, weight.
                    chart.Legend.FrameFormat.Border.AutoFormat = false;
                    chart.Legend.FrameFormat.Border.IsAutoLineColor = false;
                    chart.Legend.FrameFormat.Border.LineColor = Syncfusion.Drawing.Color.Blue;
                    chart.Legend.FrameFormat.Border.LinePattern = OfficeChartLinePattern.DashDot;
                    chart.Legend.FrameFormat.Border.LineWeight = OfficeChartLineWeight.Wide;
                    //Sets the legend's text area formatting - font name, weight, color, size.
                    chart.Legend.TextArea.Bold = true;
                    chart.Legend.TextArea.Color = OfficeKnownColors.Bright_green;
                    chart.Legend.TextArea.FontName = "Times New Roman";
                    chart.Legend.TextArea.Size = 20;
                    chart.Legend.TextArea.Strikethrough = true;
                    //Modifies the legend entry.
                    chart.Legend.LegendEntries[0].IsDeleted = true;
                    //Modifies the legend layout - height, left, top, width.
                    chart.Legend.Layout.Height = 50;
                    chart.Legend.Layout.HeightMode = LayoutModes.factor;
                    chart.Legend.Layout.Left = 10;
                    chart.Legend.Layout.LeftMode = LayoutModes.factor;
                    chart.Legend.Layout.Top = 30;
                    chart.Legend.Layout.TopMode = LayoutModes.factor;
                    chart.Legend.Layout.Width = 100;
                    chart.Legend.Layout.WidthMode = LayoutModes.factor;
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
