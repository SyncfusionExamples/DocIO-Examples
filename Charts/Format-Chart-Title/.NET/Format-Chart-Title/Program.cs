
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;


namespace Format_Chart_Title
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

                // Set the chart title.
                chart.ChartTitle = "Purchase Details";

                // Customize chart title area.
                chart.ChartTitleArea.FontName = "Calibri";
                chart.ChartTitleArea.Bold = true;
                chart.ChartTitleArea.Color = OfficeKnownColors.Black;
                chart.ChartTitleArea.Underline = OfficeUnderline.WavyHeavy;

                //Manually resizing chart title area using Layout.
                chart.ChartTitleArea.Layout.Left = 5;

                //Enable legend.
                chart.HasLegend = true;
                chart.Legend.Position = OfficeLegendPosition.Right;
               
                using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Output.docx"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
                {
                    //Save the Word file.
                    document.Save(outputStream, FormatType.Docx);
                }
            }
        }
    }
}
