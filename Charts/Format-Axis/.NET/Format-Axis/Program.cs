
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;
using System;

namespace Format_Axis
{
    class Program
    {
        static void Main(string[] args)
        {
            FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            //Open an existing document from file system through constructor of WordDocument class
            using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
            {
                //Get the paragraph
                WParagraph paragraph = document.LastParagraph;
                //Get the chart entity
                WChart chart = paragraph.ChildEntities[0] as WChart;

                //Set the horizontal (category) axis title.
                chart.PrimaryCategoryAxis.Title = "Months";
                //Set the Vertical (value) axis title.
                chart.PrimaryValueAxis.Title = "Precipitation,in.";
                //Set title for secondary value axis
                chart.SecondaryValueAxis.Title = "Temperature,deg.F";

                //Customize the horizontal category axis.
                chart.PrimaryCategoryAxis.Border.LinePattern = OfficeChartLinePattern.Solid;
                chart.PrimaryCategoryAxis.Border.LineColor = Syncfusion.Drawing.Color.Blue;
                chart.PrimaryCategoryAxis.Border.LineWeight = OfficeChartLineWeight.Hairline;

                //Customize the vertical category axis.
                chart.PrimaryValueAxis.Border.LinePattern = OfficeChartLinePattern.Solid;
                chart.PrimaryValueAxis.Border.LineColor = Syncfusion.Drawing.Color.Blue;
                chart.PrimaryValueAxis.Border.LineWeight = OfficeChartLineWeight.Narrow;

                //Customize the horizontal category axis font.
                chart.PrimaryCategoryAxis.Font.Color = OfficeKnownColors.Red;
                chart.PrimaryCategoryAxis.Font.FontName = "Calibri";
                chart.PrimaryCategoryAxis.Font.Bold = true;
                chart.PrimaryCategoryAxis.Font.Size = 8;

                //Customize the vertical category axis font.
                chart.PrimaryValueAxis.Font.Color = OfficeKnownColors.Red;
                chart.PrimaryValueAxis.Font.FontName = "Calibri";
                chart.PrimaryValueAxis.Font.Bold = true;
                chart.PrimaryValueAxis.Font.Size = 8;


                //Customize the secondary vertical category axis.
                chart.SecondaryValueAxis.Border.LinePattern = OfficeChartLinePattern.Solid;
                chart.SecondaryValueAxis.Border.LineColor = Syncfusion.Drawing.Color.Blue;
                chart.SecondaryValueAxis.Border.LineWeight = OfficeChartLineWeight.Narrow;

                //Customize the secondary vertical category axis font.
                chart.SecondaryValueAxis.Font.Color = OfficeKnownColors.Red;
                chart.SecondaryValueAxis.Font.FontName = "Calibri";
                chart.SecondaryValueAxis.Font.Bold = true;
                chart.SecondaryValueAxis.Font.Size = 8;

                //Axis title area text angle rotation.
                chart.PrimaryValueAxis.TitleArea.TextRotationAngle = 270;

                //Maximum value in the axis.
                chart.PrimaryValueAxis.MaximumValue = 15;
                chart.PrimaryValueAxis.MinimumValue = 0;
                //Number format for axis.
                chart.PrimaryValueAxis.NumberFormat = "0.0";

                //Hiding major gridlines.
                chart.PrimaryValueAxis.HasMajorGridLines = true;

                //Showing minor gridlines.
                chart.PrimaryValueAxis.HasMinorGridLines = false;

                using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Sample.docx"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
                {
                    //Save the Word file.
                    document.Save(outputStream, FormatType.Docx);
                }
                // Open the Word document located at the specified path using the default associated program.
                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo(Path.GetFullPath(@"../../../Sample.docx"))
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
