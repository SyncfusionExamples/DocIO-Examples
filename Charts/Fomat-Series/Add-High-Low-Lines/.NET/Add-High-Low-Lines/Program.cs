
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;

namespace Add_High_Low_Lines
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

                //Set HasHighLowLines property to true.
                chart.Series[0].SerieFormat.CommonSerieOptions.HasHighLowLines = true;

                //Apply formats to HighLowLines.
                chart.Series[0].SerieFormat.CommonSerieOptions.HighLowLines.LineColor = Syncfusion.Drawing.Color.Red;
                chart.Series[0].SerieFormat.CommonSerieOptions.HighLowLines.LinePattern = OfficeChartLinePattern.Solid;
                chart.Series[0].SerieFormat.CommonSerieOptions.HighLowLines.LineWeight = OfficeChartLineWeight.Medium;

                using (FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
                {
                    //Save the Word file.
                    document.Save(outputStream, FormatType.Docx);
                }
                // Open the Word document located at the specified path using the default associated program.
                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo(Path.GetFullPath(@"../../../Result.docx"))
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
