
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.OfficeChart;
using System;

namespace Add_Series_Lines
{
    class Program
    {
        static void Main(string[] args)
        {
            FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Data/Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            //Open an existing document from file system through constructor of WordDocument class.
            using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
            {
                //Get the paragraph.
                WParagraph paragraph = document.LastParagraph;
                //Get the chart entity.
                WChart chart = paragraph.ChildEntities[0] as WChart;
                for (int i = 0; i < chart.Series.Count; i++)
                {
                    //Enable the datalabel in chart.
                    chart.Series[i].DataPoints.DefaultDataPoint.DataLabels.IsValue = true;

                    // Set the font size of the data labels.
                    chart.Series[i].DataPoints.DefaultDataPoint.DataLabels.Size = 10;
                    // Change the color of the data labels. 
                    chart.Series[i].DataPoints.DefaultDataPoint.DataLabels.Color = OfficeKnownColors.Black;
                    // Make the data labels bold.
                    chart.Series[i].DataPoints.DefaultDataPoint.DataLabels.Bold = true;
                    // Set the position of data labels for the first series.
                    chart.Series[i].DataPoints.DefaultDataPoint.DataLabels.Position = OfficeDataLabelPosition.Center;

                }
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
