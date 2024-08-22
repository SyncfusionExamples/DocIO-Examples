using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Modify_an_existing_chart_data
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
                    //Gets the last paragraph.
                    WParagraph paragraph = document.LastParagraph;
                    //Gets the chart entity from the paragraph items.
                    WChart chart = paragraph.ChildEntities[0] as WChart;
                    //Modifies the data values of chart.
                    chart.ChartData.SetValue(2, 2, 200);
                    chart.ChartData.SetValue(3, 2, 20);
                    chart.ChartData.SetValue(4, 2, 20);
                    //Refreshes chart data to set the modified values.
                    chart.Refresh();
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
