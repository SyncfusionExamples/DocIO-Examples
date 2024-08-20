using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using Syncfusion.OfficeChart;
using System.IO;

namespace Applying_3D_formats
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
                chart.ChartType = OfficeChartType.Column_Clustered_3D;
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
                //Sets rotation and elevation values.
                chart.Rotation = 20;
                chart.Elevation = 15;
                //Sets side wall properties.
                chart.SideWall.Fill.FillType = OfficeFillType.SolidColor;
                chart.SideWall.Fill.ForeColor = Color.White;
                chart.SideWall.Fill.BackColor = Color.White;
                chart.SideWall.Border.LineColor = Color.Beige;
                //Sets floor fill option.
                chart.Floor.Fill.FillType = OfficeFillType.Pattern;
                //Sets the floor pattern Type.
                chart.Floor.Fill.Pattern = OfficeGradientPattern.Pat_Divot;
                //Sets the floor fore and Back ground color.
                chart.Floor.Fill.ForeColor = Color.Blue;
                chart.Floor.Fill.BackColor = Color.White;
                //Sets the floor thickness.
                chart.Floor.Thickness = 3;
                //Sets the back wall fill option.
                chart.BackWall.Fill.FillType = OfficeFillType.Gradient;
                //Sets the Texture Type.
                chart.BackWall.Fill.GradientColorType = OfficeGradientColor.TwoColor;
                chart.BackWall.Fill.GradientStyle = OfficeGradientStyle.DiagonalDown;
                chart.BackWall.Fill.ForeColor = Color.WhiteSmoke;
                chart.BackWall.Fill.BackColor = Color.LightBlue;
                //Sets the Border Line color.
                chart.BackWall.Border.LineColor = Color.Wheat;
                //Sets the Picture Type.
                chart.BackWall.PictureUnit = OfficeChartPictureType.stretch;
                //Sets the back wall thickness.
                chart.BackWall.Thickness = 10;
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
