
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;
using System;

namespace Format_Chart_Area
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

                //Adding space between bars of different series of single category.
                chart.Series[0].SerieFormat.CommonSerieOptions.Overlap = -40;

                //Adding space between bars of different categories.
                chart.Series[0].SerieFormat.CommonSerieOptions.GapWidth = 100;

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
