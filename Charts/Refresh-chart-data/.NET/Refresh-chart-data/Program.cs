using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Refresh_chart_data
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
                    // Refreshes the chart data. Set `true` to evaluate Excel formulas before refreshing,
                    // or `false` to refresh only the data without evaluating formulas.
                    chart.Refresh(false);
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
