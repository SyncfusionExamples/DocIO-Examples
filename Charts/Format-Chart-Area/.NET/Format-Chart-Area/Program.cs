
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;

namespace Format_Chart_Area
{
    class Program
    {
        static void Main(string[] args)
        {
            FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            //Open an existing document from file system through constructor of WordDocument class.
            using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
            {
                //Get the paragraph.
                WParagraph paragraph = document.LastParagraph;
                //Get the chart entity.
                WChart chart = paragraph.ChildEntities[0] as WChart;
                //Modify the chart height and width.
                chart.Height = 300;
                chart.Width = 500;

                //Format the chart area.
                IOfficeChartFrameFormat chartArea = chart.ChartArea;

                //Set border line pattern, color, line weight.
                chartArea.Border.LinePattern = OfficeChartLinePattern.Solid;
                chartArea.Border.LineColor = Syncfusion.Drawing.Color.Blue;
                chartArea.Border.LineWeight = OfficeChartLineWeight.Hairline;

                //Set fill type and fill colors.
                chartArea.Fill.FillType = OfficeFillType.Gradient;
                chartArea.Fill.GradientColorType = OfficeGradientColor.TwoColor;
                chartArea.Fill.BackColor = Syncfusion.Drawing.Color.FromArgb(205, 217, 234);
                chartArea.Fill.ForeColor = Syncfusion.Drawing.Color.White;

                using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
                {
                    //Save the Word file.
                    document.Save(outputStream, FormatType.Docx);
                }
            }
        }
    }
}
