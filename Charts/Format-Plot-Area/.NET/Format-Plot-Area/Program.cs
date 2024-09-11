
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using Syncfusion.OfficeChart;

namespace Format_Plot_Area
{
    class Program
    {
        static void Main(string[] args)
        {
            FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            //Opens an existing document from file system through constructor of WordDocument class.
            using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
            {
                //Get the paragraph.
                WParagraph paragraph = document.LastParagraph;
                //Get the chart entity.
                WChart chart = paragraph.ChildEntities[0] as WChart;
                //Modify the chart height and width.
                chart.Height = 300;
                chart.Width = 500;

                //Plot Area.
                IOfficeChartFrameFormat chartPlotArea = chart.PlotArea;

                //Plot area border settings - line pattern, color, weight.
                chartPlotArea.Border.LinePattern = OfficeChartLinePattern.Solid;
                chartPlotArea.Border.LineColor = Color.Blue;
                chartPlotArea.Border.LineWeight = OfficeChartLineWeight.Hairline;

                //Set fill type and color.
                chartPlotArea.Fill.FillType = OfficeFillType.Gradient;
                chartPlotArea.Fill.GradientColorType = OfficeGradientColor.TwoColor;
                chartPlotArea.Fill.BackColor = Color.FromArgb(205, 217, 234);
                chartPlotArea.Fill.ForeColor = Color.White;

                using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
                {
                    //Save the Word file.
                    document.Save(outputStream, FormatType.Docx);
                }
            }
        }
    }
}
